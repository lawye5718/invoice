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
    # 替换常见干扰字符，统一标点
    return text.replace(" ", "").replace("\n", "").replace("\r", "")\
               .replace("：", ":").replace("￥", "¥")\
               .replace("（", "(").replace("）", ")")\
               .replace("O", "0").replace("o", "0")

def format_date(date_str):
    """统一日期格式 YYYY-MM-DD"""
    if not date_str: return ""
    m = re.search(r'(\d{4})[-年/.](\d{1,2})[-月/.](\d{1,2})', date_str)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return ""

# ==========================================
# 2. 金额处理核心引擎 (新增大写解析)
# ==========================================

def cn_upper_to_float(cn_str):
    """
    将中文大写金额转换为浮点数
    例如：贰佰捌拾叁圆捌角壹分 -> 283.81
    """
    if not cn_str: return 0.0
    
    # 映射表
    CN_NUM = {'零': 0, '壹': 1, '贰': 2, '叁': 3, '肆': 4, '伍': 5, '陆': 6, '柒': 7, '捌': 8, '玖': 9,
              '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '两': 2}
    CN_UNIT = {'拾': 10, '十': 10, '佰': 100, '百': 100, '仟': 1000, '千': 1000, '万': 10000, '亿': 100000000}
    
    # 清洗：去掉"整"、"正"、"圆"、"元"等
    # 但要注意"圆"是分界线
    
    # 简单解析逻辑：
    # 1. 按"圆/元"分割整数和小数
    parts = re.split(r'[圆元]', cn_str)
    integer_str = parts[0]
    decimal_str = parts[1] if len(parts) > 1 else ""
    
    # --- 解析整数部分 ---
    def parse_section(s):
        val = 0
        curr_digit = 0
        unit_val = 0
        
        for char in s:
            if char in CN_NUM:
                curr_digit = CN_NUM[char]
            elif char in CN_UNIT:
                if char in ['万', '亿']:
                    # 遇到万/亿，结算前面的所有
                    val = (val + unit_val + curr_digit) * CN_UNIT[char]
                    unit_val = 0
                    curr_digit = 0
                else:
                    unit_val += curr_digit * CN_UNIT[char]
                    curr_digit = 0
            # 零忽略
        return val + unit_val + curr_digit

    total = parse_section(integer_str)
    
    # --- 解析小数部分 (角、分) ---
    decimal_val = 0.0
    curr_digit = 0
    for char in decimal_str:
        if char in CN_NUM:
            curr_digit = CN_NUM[char]
        elif char == '角':
            decimal_val += curr_digit * 0.1
            curr_digit = 0
        elif char == '分':
            decimal_val += curr_digit * 0.01
            curr_digit = 0
            
    return round(total + decimal_val, 2)

def find_amount_strict(text):
    """
    严格金额提取策略：
    1. 优先提取大写金额 (Authoritative)
    2. 提取小写金额 (Verify)
    3. 返回 (最佳金额, 备注信息)
    """
    if not text: return 0.0, "空白内容"
    
    # --- 1. 提取大写金额 ---
    # 匹配模式：价税合计(大写) XXXXX
    # 兼容：大写:、大写：
    upper_pattern = r'(?:价税合计|大写|金额).*?([零壹贰叁肆伍陆柒捌玖拾佰仟万亿圆角分整]+)'
    upper_match = re.search(upper_pattern, text)
    
    amount_upper = 0.0
    has_upper = False
    
    if upper_match:
        cn_str = upper_match.group(1)
        # 排除短杂音（例如只匹配到一个"圆"字）
        if len(cn_str) > 1:
            try:
                amount_upper = cn_upper_to_float(cn_str)
                if amount_upper > 0:
                    has_upper = True
            except: pass

    # --- 2. 提取小写金额 ---
    # 匹配模式：(小写)、¥、￥ 紧跟的数字
    # 严格模式：不再全屏搜索最大数字，防止匹配到单价
    lower_pattern = r'(?:小写|¥|￥|合计)[^0-9\.]*([0-9]{1,3}(?:,[0-9]{3})*\.[0-9]{2})'
    lower_matches = re.findall(lower_pattern, text)
    
    amount_lower = 0.0
    has_lower = False
    
    # 取第一个匹配到的有效金额 (通常发票小写就在大写后面)
    for m in lower_matches:
        try:
            val = float(m.replace(",", ""))
            if 0.01 <= val <= 5000000:
                amount_lower = val
                has_lower = True
                break # 找到即止
        except: continue

    # --- 3. 决策与校验 ---
    
    # 情况A: 有大写 (以大写为准)
    if has_upper:
        if has_lower:
            if abs(amount_upper - amount_lower) > 0.1:
                return amount_upper, f"⚠️ 大小写不符 (大写:{amount_upper} 小写:{amount_lower})"
            else:
                return amount_upper, "正常" # 校验通过
        else:
            return amount_upper, "正常 (无小写)"

    # 情况B: 无大写，有小写 (降级使用小写)
    if has_lower:
        return amount_lower, "使用小写 (未读到大写)"
        
    # 情况C: 都没有
    return 0.0, "警告:未读到金额"

def extract_seller_name_smart(text):
    """提取销售方"""
    suffix = r"[\u4e00-\u9fa5()（）]{2,30}(?:公司|事务所|酒店|旅行社|经营部|服务部|分行|支行|馆|店|处|中心)"
    candidates = list(set(re.findall(suffix, text)))
    blacklist = ["税务局", "财政部", "购买方", "开户行", "银行", "地址", "电话", "统一社会信用", "纳税人", "适用税率", "密码区", "机器编号"]
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
# 3. 解析函数 (XML & PDF)
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
        
        # XML 中的金额通常是数字，直接读取
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
            if not raw or len(raw.strip()) < 10: 
                return {
                    "发票号码": "", "开票日期": "", "销售方名称": "", "价税合计": 0.0,
                    "文件名": os.path.basename(pdf_path), "备注": "⚠️ 纯图/扫描件"
                }
            
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
            
            # 使用严格金额提取策略
            amt, status_note = find_amount_strict(text)
            seller = extract_seller_name_smart(text)
            
            # 如果没读到号但读到了金额，标注一下
            if not num and amt > 0:
                if "警告" not in status_note:
                    status_note = "无发票号-" + status_note
            
            return {
                "发票号码": num, "开票日期": date, "销售方名称": seller,
                "价税合计": amt, "文件名": os.path.basename(pdf_path),
                "备注": status_note
            }
    except: return None

# ==========================================
# 4. 校验引擎 (Verifier)
# ==========================================

class InvoiceVerifier:
    def __init__(self, processed_df):
        self.df = processed_df
        self.processed_nums = {} 
        self.processed_attrs = {}
        
        for _, row in self.df.iterrows():
            p_num = str(row.get('发票号码', '')).strip()
            p_amt = float(row.get('价税合计', 0))
            p_date = str(row.get('开票日期', '')).strip()
            p_seller = str(row.get('销售方名称', '')).strip()
            
            if p_num and len(p_num) > 6:
                self.processed_nums[p_num] = {'amount': p_amt, 'date': p_date}
            
            short_seller = p_seller[:4] if len(p_seller) >=4 else p_seller
            attr_key = (f"{p_amt:.2f}", p_date, short_seller)
            self.processed_attrs[attr_key] = True

    def check(self, raw_info):
        raw_num = str(raw_info.get('num') or raw_info.get('发票号码') or '').strip()
        raw_amt = float(raw_info.get('amount') or raw_info.get('价税合计') or 0)
        raw_date = str(raw_info.get('date') or raw_info.get('开票日期') or '').strip()
        raw_seller = str(raw_info.get('seller') or raw_info.get('销售方名称') or '').strip()
        
        # 1. 强校验：号码 + 金额
        if raw_num and len(raw_num) > 6:
            if raw_num in self.processed_nums:
                rec_amt = self.processed_nums[raw_num]['amount']
                if abs(rec_amt - raw_amt) < 0.1:
                    return True
                else:
                    return True # 号码对上了就算找到，但金额不一致是另一回事
        
        # 2. 弱校验：金额 + 日期 + 销售方 (针对无号文件)
        if raw_amt > 0:
            short_seller = raw_seller[:4] if len(raw_seller) >=4 else raw_seller
            attr_key = (f"{raw_amt:.2f}", raw_date, short_seller)
            if attr_key in self.processed_attrs:
                return True
            
            # 3. 兜底校验：金额 + 日期 (只有当金额比较独特，即有小数位时)
            if raw_amt % 1 != 0:
                for k in self.processed_attrs:
                    if k[0] == f"{raw_amt:.2f}" and k[1] == raw_date:
                        return True

        return False

# ==========================================
# 5. 主处理流程
# ==========================================

def run_process_pipeline(input_root_dir, output_dir):
    merged_dir = os.path.join(output_dir, 'Merged_PDFs')
    noxml_dir = os.path.join(output_dir, 'No_XML_PDFs')
    os.makedirs(merged_dir, exist_ok=True)
    os.makedirs(noxml_dir, exist_ok=True)

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
                amt, _ = find_amount_strict(text) # 使用严格模式
                folder = os.path.dirname(pdf)
                if is_trip_file(os.path.basename(pdf), text):
                    trip_pool.append({'path': pdf, 'amount': amt, 'folder': folder, 'used': False})
                else:
                    invoice_pdf_pool.append({'path': pdf, 'amount': amt, 'folder': folder})
        except: pass

    excel_rows = []
    idx = 1
    processed_source_files = set()

    # A. XML处理
    for xml in xml_files:
        info = parse_xml_invoice_data(xml)
        if not info: continue
        processed_source_files.add(os.path.abspath(xml))
        
        row = {"序号": idx, "发票号码": info['num'], "开票日期": info['date'],
               "销售方名称": info['seller'], "价税合计": info['amount'], "数据来源": "XML", "文件名": os.path.basename(xml), "备注": "正常"}
        
        folder = os.path.dirname(xml)
        target_pdf = None
        cands = [p['path'] for p in invoice_pdf_pool if p['folder'] == folder]
        xml_base = os.path.splitext(os.path.basename(xml))[0]
        
        for p in cands:
            if xml_base in os.path.basename(p) or (info['num'] and info['num'] in os.path.basename(p)):
                target_pdf = p; break
        
        if target_pdf:
            processed_source_files.add(os.path.abspath(target_pdf))
            matched_trip = None
            for t in [x for x in trip_pool if x['folder'] == folder and not x['used']]:
                if abs(t['amount'] - info['amount']) < 0.05:
                    matched_trip = t; t['used'] = True; break
            
            if matched_trip:
                processed_source_files.add(os.path.abspath(matched_trip['path']))
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
        else:
             row['备注'] = "仅XML(缺PDF)"
        
        excel_rows.append(row); idx += 1

    # B. PDF处理
    for inv in invoice_pdf_pool:
        if os.path.abspath(inv['path']) in processed_source_files: continue
        
        data = extract_data_from_pdf_simple(inv['path'])
        if not data: continue
        processed_source_files.add(os.path.abspath(inv['path']))
        
        matched_trip = None
        folder = inv['folder']
        for t in [x for x in trip_pool if x['folder'] == folder and not x['used']]:
            if inv['amount'] > 0 and abs(t['amount'] - inv['amount']) < 0.05:
                matched_trip = t; t['used'] = True; break
        
        if matched_trip:
            processed_source_files.add(os.path.abspath(matched_trip['path']))
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
            
        data['序号'] = idx; excel_rows.append(data); idx += 1

    # C. 行程单兜底
    for t in trip_pool:
        if not t['used']:
            processed_source_files.add(os.path.abspath(t['path']))
            try: shutil.copy2(t['path'], os.path.join(noxml_dir, os.path.basename(t['path'])))
            except: pass

    # D. 生成结果与核对
    excel_path = None
    df_result = pd.DataFrame()
    if excel_rows:
        df_result = pd.DataFrame(excel_rows)
        cols = ["序号", "发票号码", "开票日期", "销售方名称", "价税合计", "数据来源", "备注", "文件名"]
        for c in cols: 
            if c not in df_result.columns: df_result[c] = ""
        df_result = df_result[cols]
        df_result['价税合计'] = pd.to_numeric(df_result['价税合计'], errors='coerce').fillna(0.0)
        
        sum_row = {"序号": "总计", "价税合计": df_result['价税合计'].sum(), "销售方名称": f"共 {len(df_result)} 张"}
        df_display = pd.concat([df_result, pd.DataFrame([sum_row])], ignore_index=True)
        excel_path = os.path.join(output_dir, 'Summary_Final.xlsx')
        df_display.to_excel(excel_path, index=False)

    missing_files = []
    if not df_result.empty:
        verifier = InvoiceVerifier(df_result)
        for f in all_files:
            if not f.lower().endswith(('.pdf', '.xml')): continue
            
            raw_info = None
            try:
                if f.lower().endswith('.xml'): raw_info = parse_xml_invoice_data(f)
                else: raw_info = extract_data_from_pdf_simple(f)
            except: pass
            
            is_missing = True
            if raw_info:
                # 特殊处理：如果PDF被识别为纯图(备注含警告)，直接算遗漏
                if "纯图" in raw_info.get('备注', ''):
                    is_missing = True
                elif verifier.check(raw_info):
                    is_missing = False
            
            if is_missing:
                missing_files.append(f)

    return excel_path, merged_dir, noxml_dir, missing_files

# ==========================================
# 6. 手动核对功能
# ==========================================

def run_manual_check(raw_dir, proc_zip_path, out_dir):
    df_proc = pd.DataFrame()
    with zipfile.ZipFile(proc_zip_path, 'r') as z:
        xls = [n for n in z.namelist() if n.endswith('.xlsx')]
        if xls:
            with z.open(xls[0]) as f: df_proc = pd.read_excel(f)
        else:
            # 兼容旧版ZIP
            rows = []
            for n in z.namelist():
                if n.endswith('.pdf'):
                    base = os.path.basename(n)
                    m = re.match(r'(\d+)_([\d\.]+)\.pdf', base)
                    if m: rows.append({'发票号码': m.group(1), '价税合计': float(m.group(2))})
            df_proc = pd.DataFrame(rows)

    verifier = InvoiceVerifier(df_proc)
    missing = []
    matched_count = 0
    
    for root, _, files in os.walk(raw_dir):
        for f in files:
            if not f.lower().endswith(('.pdf', '.xml')): continue
            fp = os.path.join(root, f)
            
            raw_info = None
            try:
                if f.lower().endswith('.xml'): raw_info = parse_xml_invoice_data(fp)
                else: raw_info = extract_data_from_pdf_simple(fp)
            except: pass
            
            if raw_info and verifier.check(raw_info):
                matched_count += 1
            else:
                missing.append(fp)
    
    zip_p = None
    if missing:
        zip_p = os.path.join(out_dir, "Manual_Missing.zip")
        with zipfile.ZipFile(zip_p, 'w', zipfile.ZIP_DEFLATED) as z:
            for m in missing: z.write(m, os.path.basename(m))
            
    return matched_count, len(missing), zip_p

# ==========================================
# 7. Streamlit 主界面
# ==========================================

def main():
    st.set_page_config(page_title="发票无忧 V12 (大写金额校验版)", layout="wide")
    st.title("🧾 发票无忧 V12 (严格金额校验)")

    tab1, tab2 = st.tabs(["🚀 一键处理", "🔍 手动复核"])

    with tab1:
        st.info("上传 ZIP/文件夹 -> 优先提取大写金额 -> 自动核对")
        uploaded_files = st.file_uploader("上传文件", type=['zip', 'xml', 'pdf'], accept_multiple_files=True, key="u1")

        if uploaded_files and st.button("开始处理", key="b1"):
            with st.spinner('正在处理...'):
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
                    
                    st.success("✅ 完成！")
                    c1, c2 = st.columns(2)
                    if excel:
                        df = pd.read_excel(excel)
                        c1.metric("已录入", f"{len(df)-1} 张")
                        st.dataframe(df.tail(3))
                        res_zip = os.path.join(temp_dir, "Result.zip")
                        with zipfile.ZipFile(res_zip, 'w', zipfile.ZIP_DEFLATED) as z:
                            z.write(excel, "汇总表.xlsx")
                            for r, _, fs in os.walk(merged):
                                for f in fs: z.write(os.path.join(r, f), f"合并后发票/{f}")
                            for r, _, fs in os.walk(noxml):
                                for f in fs: z.write(os.path.join(r, f), f"独立发票/{f}")
                        with open(res_zip, "rb") as f:
                            st.download_button("📥 下载结果 (Result.zip)", f, "Result.zip")

                    c2.metric("遗漏", f"{len(missing_list)} 个", delta_color="inverse")
                    if missing_list:
                        m_zip = os.path.join(temp_dir, "Missing.zip")
                        with zipfile.ZipFile(m_zip, 'w', zipfile.ZIP_DEFLATED) as z:
                            for m in missing_list: z.write(m, f"遗漏文件/{os.path.basename(m)}")
                        with open(m_zip, "rb") as f:
                            st.download_button("📥 下载遗漏包 (Missing.zip)", f, "Missing.zip", type="primary")

    with tab2:
        st.write("反向核对：用 Excel 结果检查原始文件是否遗漏。")
        c1, c2 = st.columns(2)
        raw_ups = c1.file_uploader("1. 上传原始文件", type=['zip','pdf'], accept_multiple_files=True, key="u2")
        proc_zip = c2.file_uploader("2. 上传 Result.zip", type=['zip'], key="u3")
        
        if raw_ups and proc_zip and st.button("开始核对", key="b2"):
            with st.spinner("核对中..."):
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
                    st.metric("❌ 遗漏", miss)
                    if mzip:
                        with open(mzip, "rb") as f:
                            st.download_button("📥 下载遗漏文件", f, "Manual_Missing.zip")

if __name__ == "__main__":
    main()