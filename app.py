import streamlit as st
import os
import zipfile
import shutil
import tempfile
import re
import pandas as pd
import pdfplumber
import xml.etree.ElementTree as ET
from pypdf import PdfWriter

# ==========================================
# 1. 基础工具函数
# ==========================================

def extract_zip_with_encoding(zip_path, extract_to):
    """解压 ZIP 并修复中文乱码"""
    with zipfile.ZipFile(zip_path, 'r') as z:
        for file_info in z.infolist():
            try:
                if file_info.flag_bits & 0x800 == 0:
                    original_name = file_info.filename.encode('cp437').decode('gbk')
                else:
                    original_name = file_info.filename
            except:
                try: original_name = file_info.filename.encode('utf-8').decode('utf-8')
                except: original_name = file_info.filename

            if "__MACOSX" in original_name or ".DS_Store" in original_name:
                continue

            target_path = os.path.join(extract_to, original_name)
            parent_dir = os.path.dirname(target_path)
            if parent_dir and not os.path.exists(parent_dir):
                os.makedirs(parent_dir, exist_ok=True)
                
            if not original_name.endswith('/'):
                with z.open(file_info) as source, open(target_path, "wb") as target:
                    shutil.copyfileobj(source, target)

def normalize_text(text):
    """文本清洗"""
    if not text: return ""
    return text.replace(" ", "").replace("\n", "").replace("\r", "")\
               .replace("：", ":").replace("￥", "¥")\
               .replace("（", "(").replace("）", ")")\
               .replace("O", "0")

def format_date(date_str):
    """统一日期格式 YYYY-MM-DD"""
    if not date_str: return ""
    # 提取年-月-日
    m = re.search(r'(\d{4})[-年/.](\d{1,2})[-月/.](\d{1,2})', date_str)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return ""

def find_best_amount(text):
    """双重策略提取金额"""
    if not text: return 0.0
    # 策略A: 锚点查找
    anchor_pattern = r'(?:小写|¥|￥|合计|金额)[^0-9\.]*([0-9]{1,3}(?:,[0-9]{3})*\.[0-9]{2})'
    for m in re.findall(anchor_pattern, text):
        try:
            val = float(m.replace(",", ""))
            if 0.01 <= val <= 5000000: return val
        except: continue

    # 策略B: 最大值
    matches = re.findall(r'(\d{1,3}(?:,\d{3})*\.\d{2})', text)
    valid = []
    for m in matches:
        try:
            val = float(m.replace(",", ""))
            if 0.01 <= val <= 5000000 and val not in [0.06, 0.03, 0.13, 0.01, 1.00]:
                valid.append(val)
        except: continue
    return max(valid) if valid else 0.0

def extract_seller_name_smart(text):
    """提取销售方"""
    suffix = r"[\u4e00-\u9fa5()（）]{2,30}(?:公司|事务所|酒店|旅行社|经营部|服务部|分行|支行|馆|店|处|中心)"
    candidates = list(set(re.findall(suffix, text)))
    blacklist = ["税务局", "财政部", "购买方", "开户行", "银行", "地址", "电话", "统一社会信用", "纳税人", "适用税率", "密码区"]
    filtered = [c for c in candidates if not any(b in c for b in blacklist) and len(c) >= 4]
    return max(filtered, key=len) if filtered else ""

def is_trip_file(filename, text=None):
    """判断行程单"""
    fn = filename.lower()
    if "行程" in fn or "trip" in fn or "报销" in fn:
        if text:
            clean = normalize_text(text)
            if "发票代码" in clean or "发票号码" in clean or "电子发票" in clean:
                return False
        return True
    return False

# ==========================================
# 2. 解析函数 (增强版)
# ==========================================

def parse_xml_invoice_data(xml_path):
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        def g(path):
            node = root.find(path)
            return node.text if node is not None else ""

        num = g(".//TaxSupervisionInfo/InvoiceNumber") or g(".//InvoiceNumber") or g(".//Fphm")
        date = g(".//TaxSupervisionInfo/IssueTime") or g(".//IssueTime") or g(".//Kprq")
        seller = g(".//SellerInformation/SellerName") or g(".//Xfmc")
        amt_str = g(".//BasicInformation/TotalTax-includedAmount") or g(".//TotalTax-includedAmount") or g(".//TotalAmount") or g(".//Jshj")
        amount = float(amt_str.replace(',', '')) if amt_str else 0.0

        return {
            "num": num, "date": format_date(date), 
            "seller": seller, "amount": amount
        }
    except: return None

def extract_data_from_pdf_simple(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as p:
            if not p.pages: return None
            raw = p.pages[0].extract_text()
            if not raw or len(raw.strip()) < 10: return None # 纯图忽略
            
            text = normalize_text(raw)
            num = ""
            m20 = re.search(r'(\d{20})', text)
            if m20: num = m20.group(1)
            else:
                m8 = re.search(r'(?:号码|No)[:|]?(\d{8,})', text)
                if m8: num = m8.group(1)
            
            date = ""
            md = re.search(r'(\d{4}[-年/.]\d{1,2}[-月/.]\d{1,2}日?)', text)
            if md: date = format_date(md.group(1))
            
            amt = find_best_amount(text)
            seller = extract_seller_name_smart(text)
            
            return {
                "发票号码": num, "开票日期": date, "销售方名称": seller,
                "价税合计": amt, "文件名": os.path.basename(pdf_path)
            }
    except: return None

# ==========================================
# 3. 核心校验引擎 (Verification Engine)
# ==========================================

class InvoiceVerifier:
    def __init__(self, processed_df):
        """
        根据 Excel 数据建立多维索引
        """
        self.df = processed_df
        self.processed_nums = {}   # Num -> {Amount, Date}
        self.processed_attrs = {}  # (Amount, Date, Seller) -> Exists
        
        # 建立索引
        for _, row in self.df.iterrows():
            # 1. 索引号码
            p_num = str(row.get('发票号码', '')).strip()
            p_amt = float(row.get('价税合计', 0))
            p_date = str(row.get('开票日期', '')).strip()
            p_seller = str(row.get('销售方名称', '')).strip()
            
            if p_num and len(p_num) > 6: # 忽略太短的号码
                self.processed_nums[p_num] = {'amount': p_amt, 'date': p_date}
            
            # 2. 索引属性 (金额+日期+销售方) - 用于无号匹配
            # 销售方只取前4个字作为模糊匹配键，防止公司名全称/简称差异
            short_seller = p_seller[:4] if len(p_seller) >=4 else p_seller
            attr_key = (f"{p_amt:.2f}", p_date, short_seller)
            self.processed_attrs[attr_key] = True

    def check(self, raw_info):
        """
        核对原始文件是否在已处理列表中
        返回: (是否通过, 原因)
        """
        raw_num = str(raw_info.get('num') or '').strip()
        raw_amt = float(raw_info.get('amount') or 0)
        raw_date = str(raw_info.get('date') or '').strip()
        raw_seller = str(raw_info.get('seller') or '').strip()
        
        # 1. 优先匹配发票号码 (强校验)
        if raw_num and len(raw_num) > 6:
            if raw_num in self.processed_nums:
                # 进一步核对金额 (允许 0.1 误差)
                rec_amt = self.processed_nums[raw_num]['amount']
                if abs(rec_amt - raw_amt) < 0.1:
                    return True, "号码与金额完全匹配"
                else:
                    # 号码对但金额不对，依然算"已处理"，但值得注意
                    # 在核对"遗漏"的语境下，只要Excel里有这个号，就不算遗漏
                    return True, f"号码匹配但金额不一致 (Excel:{rec_amt} vs Raw:{raw_amt})"
        
        # 2. 如果没有号码或号码没匹配上，尝试"无号匹配" (兜底)
        # 只有当金额 > 0 时才进行此匹配
        if raw_amt > 0:
            short_seller = raw_seller[:4] if len(raw_seller) >=4 else raw_seller
            attr_key = (f"{raw_amt:.2f}", raw_date, short_seller)
            if attr_key in self.processed_attrs:
                return True, "金额、日期、销售方匹配"
                
            # 放宽条件：只匹配 金额 + 日期 (防止销售方识别失败)
            # 但为了防止撞车，只有当金额比较"独特"(带小数)时才敢认
            if raw_amt % 1 != 0:
                for k in self.processed_attrs:
                    # k = (amt_str, date, seller)
                    if k[0] == f"{raw_amt:.2f}" and k[1] == raw_date:
                        return True, "金额与日期匹配(忽略销售方)"

        # 3. 确实找不到
        return False, "未找到匹配项"

# ==========================================
# 4. 主处理流程
# ==========================================

def run_process_pipeline(input_root_dir, output_dir):
    """处理并自动核对"""
    merged_dir = os.path.join(output_dir, 'Merged_PDFs')
    noxml_dir = os.path.join(output_dir, 'No_XML_PDFs')
    os.makedirs(merged_dir, exist_ok=True)
    os.makedirs(noxml_dir, exist_ok=True)

    # 1. 扫描与池化
    all_files = []
    for root, dirs, files in os.walk(input_root_dir):
        for f in files: all_files.append(os.path.join(root, f))
    
    xml_files = [f for f in all_files if f.lower().endswith('.xml')]
    pdf_files = [f for f in all_files if f.lower().endswith('.pdf')]
    
    trip_pool = []
    invoice_pdf_pool = []
    
    for pdf in pdf_files:
        try:
            with pdfplumber.open(pdf) as p:
                if not p.pages: continue
                text = normalize_text(p.pages[0].extract_text())
                amt = find_best_amount(text)
                folder = os.path.dirname(pdf)
                if is_trip_file(os.path.basename(pdf), text):
                    trip_pool.append({'path': pdf, 'amount': amt, 'folder': folder, 'used': False})
                else:
                    invoice_pdf_pool.append({'path': pdf, 'amount': amt, 'folder': folder})
        except: pass

    # 2. 处理流程
    excel_rows = []
    idx = 1
    processed_files_map = set() # 仅用于去重，不用于核对

    # A. XML处理
    for xml in xml_files:
        info = parse_xml_invoice_data(xml)
        if not info: continue
        processed_files_map.add(os.path.abspath(xml))
        
        row = {"序号": idx, "发票号码": info['num'], "开票日期": info['date'],
               "销售方名称": info['seller'], "价税合计": info['amount'], "数据来源": "XML", "文件名": os.path.basename(xml)}
        
        folder = os.path.dirname(xml)
        target_pdf = None
        # 找PDF
        cands = [p['path'] for p in invoice_pdf_pool if p['folder'] == folder]
        for p in cands:
            if os.path.splitext(os.path.basename(xml))[0] in os.path.basename(p) or (info['num'] and info['num'] in os.path.basename(p)):
                target_pdf = p; break
        
        if target_pdf:
            processed_files_map.add(os.path.abspath(target_pdf))
            # 找行程单
            matched_trip = None
            for t in [x for x in trip_pool if x['folder'] == folder and not x['used']]:
                if abs(t['amount'] - info['amount']) < 0.05:
                    matched_trip = t; t['used'] = True; break
            
            if matched_trip:
                try:
                    merger = PdfWriter()
                    merger.append(target_pdf); merger.append(matched_trip['path'])
                    safe_name = f"{info['num']}_{info['amount']}.pdf".replace(':','').replace('/','_')
                    merger.write(os.path.join(merged_dir, safe_name)); merger.close()
                    row['备注'] = "已合并行程单"
                except:
                    shutil.copy2(target_pdf, os.path.join(noxml_dir, os.path.basename(target_pdf)))
                    row['备注'] = "合并失败-保留原件"
            else:
                shutil.copy2(target_pdf, os.path.join(noxml_dir, os.path.basename(target_pdf)))
                row['备注'] = "正常(无行程单)"
        else:
             row['备注'] = "仅XML(缺PDF)"
        
        excel_rows.append(row); idx += 1

    # B. PDF处理
    for inv in invoice_pdf_pool:
        if os.path.abspath(inv['path']) in processed_files_map: continue
        data = extract_data_from_pdf_simple(inv['path'])
        if not data: continue # 无法识别的文件暂不进表，依靠核对环节捞回
        
        # 找行程单
        matched_trip = None
        folder = inv['folder']
        for t in [x for x in trip_pool if x['folder'] == folder and not x['used']]:
            if inv['amount'] > 0 and abs(t['amount'] - inv['amount']) < 0.05:
                matched_trip = t; t['used'] = True; break
        
        if matched_trip:
            try:
                merger = PdfWriter()
                merger.append(inv['path']); merger.append(matched_trip['path'])
                num = data.get('发票号码', 'NoNum')
                safe_name = f"{num}_{inv['amount']}.pdf".replace(':','').replace('/','_')
                merger.write(os.path.join(merged_dir, safe_name)); merger.close()
                data['备注'] = "已合并行程单"
                if data['价税合计'] == 0: data['价税合计'] = inv['amount']
            except:
                shutil.copy2(inv['path'], os.path.join(noxml_dir, os.path.basename(inv['path'])))
                data['备注'] = "合并失败-保留原件"
        else:
            shutil.copy2(inv['path'], os.path.join(noxml_dir, os.path.basename(inv['path'])))
            data['备注'] = "正常(无XML)"
            
        data['序号'] = idx; excel_rows.append(data); idx += 1

    # C. 生成 Excel
    excel_path = None
    df_result = pd.DataFrame()
    if excel_rows:
        df_result = pd.DataFrame(excel_rows)
        cols = ["序号", "发票号码", "开票日期", "销售方名称", "价税合计", "数据来源", "备注", "文件名"]
        for c in cols: 
            if c not in df_result.columns: df_result[c] = ""
        df_result = df_result[cols]
        df_result['价税合计'] = pd.to_numeric(df_result['价税合计'], errors='coerce').fillna(0.0)
        
        # 保存用于核对
        df_final = df_result.copy() 
        
        # 添加总计行 (只为了展示，核对时不包含)
        sum_row = {"序号": "总计", "价税合计": df_result['价税合计'].sum(), "销售方名称": f"共 {len(df_result)} 张"}
        df_display = pd.concat([df_result, pd.DataFrame([sum_row])], ignore_index=True)
        excel_path = os.path.join(output_dir, 'Summary_Final.xlsx')
        df_display.to_excel(excel_path, index=False)
    
    # 3. --- 自动核对环节 (Auto Verification) ---
    missing_files = []
    
    if not df_result.empty:
        verifier = InvoiceVerifier(df_result) # 基于结果建立索引
        
        # 遍历所有原始文件进行核对
        for f in all_files:
            if not f.lower().endswith(('.pdf', '.xml')): continue
            
            # 提取原始文件特征
            raw_info = {}
            try:
                if f.lower().endswith('.xml'):
                    raw_info = parse_xml_invoice_data(f)
                else:
                    # 对于PDF，如果它是被用掉的行程单，我们暂时不视为遗漏
                    # 但为了严谨，我们检查它是否在 Excel 或被标记为行程单
                    # 简化逻辑：尝试作为发票提取
                    raw_info = extract_data_from_pdf_simple(f)
                    # 如果提取失败(如扫描件)，raw_info 为 None
            except: pass
            
            # 判定逻辑
            is_missing = False
            
            if not raw_info:
                # 无法解析的文件，视为遗漏 (可能是坏文件或纯图)
                is_missing = True
            else:
                # 检查是否在结果中
                found, reason = verifier.check({
                    'num': raw_info.get('num') or raw_info.get('发票号码'),
                    'amount': raw_info.get('amount') or raw_info.get('价税合计'),
                    'date': raw_info.get('date') or raw_info.get('开票日期'),
                    'seller': raw_info.get('seller') or raw_info.get('销售方名称')
                })
                
                if not found:
                    # 如果没找到，再给一次机会：是不是行程单？
                    # 如果是行程单，且在 Excel 备注里有 "已合并行程单" 的记录，我们很难一一对应
                    # 所以策略是：只报告 "遗漏的发票"。行程单如果没匹配上，也是一种遗漏。
                    # 这里直接判定为遗漏。
                     is_missing = True
            
            if is_missing:
                # 排除掉确实是行程单且被程序内部消化的情况？
                # 现在的逻辑更严格：只要 Excel 里找不到这个号/金额，就算遗漏。
                # 这会把未合并的行程单也算作遗漏（这是好事，用户需要知道哪些行程单没用上）
                missing_files.append(f)

    return excel_path, merged_dir, noxml_dir, missing_files

# ==========================================
# 5. 手动核对 (复用 Verifier)
# ==========================================

def run_manual_check(raw_dir, proc_zip_path, out_dir):
    # 1. 解压并读取 Excel
    df_proc = pd.DataFrame()
    with zipfile.ZipFile(proc_zip_path, 'r') as z:
        # 优先找 Excel
        xls = [n for n in z.namelist() if n.endswith('.xlsx')]
        if xls:
            with z.open(xls[0]) as f:
                df_proc = pd.read_excel(f)
        else:
            # 没有 Excel，回退到文件名解析 (旧逻辑兼容)
            rows = []
            for n in z.namelist():
                if n.endswith('.pdf'):
                    base = os.path.basename(n)
                    # 尝试从文件名提取 num, amount
                    m = re.match(r'(\d+)_([\d\.]+)\.pdf', base)
                    if m:
                        rows.append({'发票号码': m.group(1), '价税合计': float(m.group(2))})
            df_proc = pd.DataFrame(rows)

    verifier = InvoiceVerifier(df_proc)
    
    # 2. 遍历原始文件
    missing = []
    matched_count = 0
    
    for root, _, files in os.walk(raw_dir):
        for f in files:
            if not f.lower().endswith(('.pdf', '.xml')): continue
            fp = os.path.join(root, f)
            
            raw_info = {}
            try:
                if f.lower().endsWith('.xml'): raw_info = parse_xml_invoice_data(fp)
                else: raw_info = extract_data_from_pdf_simple(fp)
            except: pass
            
            if not raw_info:
                missing.append(fp)
                continue
                
            found, _ = verifier.check({
                'num': raw_info.get('num') or raw_info.get('发票号码'),
                'amount': raw_info.get('amount') or raw_info.get('价税合计'),
                'date': raw_info.get('date') or raw_info.get('开票日期'),
                'seller': raw_info.get('seller') or raw_info.get('销售方名称')
            })
            
            if found: matched_count += 1
            else: missing.append(fp)
            
    # 打包
    zip_p = None
    if missing:
        zip_p = os.path.join(out_dir, "Manual_Missing.zip")
        with zipfile.ZipFile(zip_p, 'w', zipfile.ZIP_DEFLATED) as z:
            for m in missing: z.write(m, os.path.basename(m))
            
    return matched_count, len(missing), zip_p

# ==========================================
# 6. Streamlit 主界面
# ==========================================

def main():
    st.set_page_config(page_title="发票无忧 V11 (精准核对版)", layout="wide")
    st.title("🧾 发票无忧 V11 (高精度自动核对)")

    tab1, tab2 = st.tabs(["🚀 一键处理+核对", "🔍 手动复核"])

    # --- Tab 1 ---
    with tab1:
        st.info("上传 ZIP/文件夹 -> 处理 -> 自动比对结果 -> 导出遗漏文件")
        uploaded_files = st.file_uploader("上传文件", type=['zip', 'xml', 'pdf'], accept_multiple_files=True, key="u1")

        if uploaded_files and st.button("开始处理", key="b1"):
            with st.spinner('正在处理并进行全量核对...'):
                with tempfile.TemporaryDirectory() as temp_dir:
                    input_root = os.path.join(temp_dir, "input")
                    os.makedirs(input_root, exist_ok=True)
                    
                    for i, up in enumerate(uploaded_files):
                        scope_dir = os.path.join(input_root, f"scope_{i}")
                        os.makedirs(scope_dir, exist_ok=True)
                        save_path = os.path.join(scope_dir, up.name)
                        with open(save_path, "wb") as f: f.write(up.getbuffer())
                        if up.name.endswith('.zip'):
                            extract_zip_with_encoding(save_path, scope_dir)
                            os.remove(save_path)
                    
                    out_dir = os.path.join(temp_dir, "output")
                    excel, merged, noxml, missing_list = run_process_pipeline(input_root, out_dir)
                    
                    st.success("✅ 处理完成！")
                    c1, c2 = st.columns(2)
                    
                    if excel:
                        df = pd.read_excel(excel)
                        c1.metric("已录入发票", f"{len(df)-1} 张")
                        st.dataframe(df.tail(3))
                        
                        # 下载结果
                        res_zip = os.path.join(temp_dir, "Result.zip")
                        with zipfile.ZipFile(res_zip, 'w', zipfile.ZIP_DEFLATED) as z:
                            z.write(excel, "汇总表.xlsx")
                            for r, _, fs in os.walk(merged):
                                for f in fs: z.write(os.path.join(r, f), f"合并后发票/{f}")
                            for r, _, fs in os.walk(noxml):
                                for f in fs: z.write(os.path.join(r, f), f"独立发票/{f}")
                        with open(res_zip, "rb") as f:
                            st.download_button("📥 下载结果包 (Result.zip)", f, "Result.zip")

                    # 遗漏报告
                    c2.metric("遗漏文件 (含无效/未匹配)", f"{len(missing_list)} 个", delta_color="inverse")
                    if missing_list:
                        st.error("检测到遗漏文件！(已打包，包含坏文件、扫描件或未匹配的行程单)")
                        m_zip = os.path.join(temp_dir, "Missing.zip")
                        with zipfile.ZipFile(m_zip, 'w', zipfile.ZIP_DEFLATED) as z:
                            for m in missing_list: z.write(m, f"遗漏文件/{os.path.basename(m)}")
                        with open(m_zip, "rb") as f:
                            st.download_button("📥 下载遗漏包 (Missing.zip)", f, "Missing.zip", type="primary")

    # --- Tab 2 ---
    with tab2:
        st.write("用【Excel 数据】反向核对原始文件，精度更高。")
        c1, c2 = st.columns(2)
        raw_ups = c1.file_uploader("1. 上传原始发票", type=['zip','pdf'], accept_multiple_files=True, key="u2")
        proc_zip = c2.file_uploader("2. 上传 Result.zip", type=['zip'], key="u3")
        
        if raw_ups and proc_zip and st.button("开始核对", key="b2"):
            with st.spinner("正在解压比对..."):
                with tempfile.TemporaryDirectory() as td:
                    raw_d = os.path.join(td, "raw")
                    os.makedirs(raw_d, exist_ok=True)
                    for up in raw_ups:
                        p = os.path.join(raw_d, up.name)
                        with open(p, "wb") as f: f.write(up.getbuffer())
                        if p.endswith('.zip'): extract_zip_with_encoding(p, raw_d)
                    
                    pz = os.path.join(td, "proc.zip")
                    with open(pz, "wb") as f: f.write(proc_zip.getbuffer())
                    
                    match, miss, mzip = run_manual_check(raw_d, pz, td)
                    st.metric("✅ 匹配成功", match)
                    st.metric("❌ 遗漏/未录入", miss)
                    if mzip:
                        with open(mzip, "rb") as f:
                            st.download_button("📥 下载遗漏文件", f, "Manual_Missing.zip")

if __name__ == "__main__":
    main()