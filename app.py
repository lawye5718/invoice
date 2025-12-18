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
from difflib import SequenceMatcher

# ==========================================
# 1. 基础工具函数
# ==========================================

def extract_zip_with_encoding(zip_path, extract_to):
    """解压 ZIP 并修复中文乱码"""
    with zipfile.ZipFile(zip_path, 'r') as z:
        for file_info in z.infolist():
            try:
                # 尝试修复文件名编码 (CP437 -> GBK)
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
    m = re.search(r'(\d{4})[-年/.](\d{1,2})[-月/.](\d{1,2})', date_str)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return ""

# ==========================================
# 2. 核心逻辑：文件类型判定与匹配
# ==========================================

def is_trip_file(filename, text=None):
    """
    判断是否为行程单/报销单
    """
    fn = filename.lower()
    
    # 1. 文件名特征 (增加 '报销单')
    if "行程" in fn or "trip" in fn or "报销" in fn:
        if text:
            clean = normalize_text(text)
            # 防误判逻辑：只有当出现明确的"发票号码+数字"时，才认为它是发票
            # 仅有"电子发票"四个字（如免责声明）不作为排除依据
            if re.search(r'发票号码[:|]?\d{8}', clean):
                return False
            # 某些行程单虽然含"发票"字样，但如果有"行程单"或"Trip"字样且无发票号，仍视为行程单
        return True
        
    # 2. 内容特征兜底
    if text:
        clean = normalize_text(text)
        if ("行程单" in clean or "triptable" in clean) and not re.search(r'发票号码[:|]?\d{8}', clean):
             return True

    return False

def clean_filename_for_matching(filename):
    """清洗文件名，提取核心特征"""
    name = os.path.splitext(filename)[0]
    # 去除通用无意义词汇，保留核心标识(如订单号、人名)
    keywords = [
        "电子发票", "普通发票", "发票", "invoice", 
        "行程报销单", "报销单", "行程单", "行程", "trip", "travel",
        "滴滴", "出行", "客票", "航空", "机票",
        "copy", "副本", "下载", "download"
    ]
    for k in keywords:
        name = name.replace(k, "")
    # 去除符号
    name = re.sub(r'[ _\-\(\)（）]', "", name)
    return name.lower()

def is_filename_match(name1, name2):
    """判断两个文件名是否高度相关"""
    c1 = clean_filename_for_matching(name1)
    c2 = clean_filename_for_matching(name2)
    if not c1 or not c2: return False
    # 包含关系或高度相似
    if c1 in c2 or c2 in c1: return True
    return SequenceMatcher(None, c1, c2).ratio() > 0.85

def get_matching_trip_advanced(invoice_amount, invoice_filename, folder, trip_pool):
    """
    增强匹配引擎：文件名匹配优先 -> 金额校验
    """
    # 筛选同 Scope (同文件夹) 下未使用的行程单
    candidates = [t for t in trip_pool if t['folder'] == folder and not t['used']]
    if not candidates: return None, None

    # --- 策略 A: 优先文件名匹配 (符合用户"一一匹配"的描述) ---
    for t in candidates:
        if is_filename_match(invoice_filename, os.path.basename(t['path'])):
            # 找到文件名匹配的，立即校验金额
            trip_amt = t['amount']
            inv_amt = invoice_amount if invoice_amount > 0 else 0
            
            # 金额校验 (允许 0.1 元误差)
            if inv_amt > 0 and trip_amt > 0:
                if abs(trip_amt - inv_amt) < 0.1:
                    return t, "正常(文件名+金额匹配)"
                else:
                    # 金额不符，但文件名匹配，依然合并，但标记需复核
                    return t, f"❌ 金额不符(发票:{inv_amt} vs 行程:{trip_amt}) 需人工复核"
            else:
                # 其中一方没读到金额，但文件名对了，也合并
                return t, "文件名匹配-金额缺失(需核对)"

    # --- 策略 B: 金额精准匹配 (作为补充) ---
    # 如果文件名没对上（可能重命名了），但金额完全一致，也认
    if invoice_amount > 0:
        for t in candidates:
            if abs(t['amount'] - invoice_amount) < 0.05:
                return t, "金额匹配(文件名不符)"
    
    # --- 策略 C: 同包唯一性兜底 ---
    # 一个包里剩最后一张发票和一张行程单
    if len(candidates) == 1:
        return candidates[0], "唯一匹配(慎用-需核对)"

    return None, None

# ==========================================
# 3. 数据提取 (含大写解析)
# ==========================================

def cn_upper_to_float(cn_str):
    """中文大写金额转数字"""
    if not cn_str: return 0.0
    CN_NUM = {'零': 0, '壹': 1, '贰': 2, '叁': 3, '肆': 4, '伍': 5, '陆': 6, '柒': 7, '捌': 8, '玖': 9,
              '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '两': 2}
    CN_UNIT = {'拾': 10, '十': 10, '佰': 100, '百': 100, '仟': 1000, '千': 1000, '万': 10000, '亿': 100000000}
    parts = re.split(r'[圆元]', cn_str)
    integer_str = parts[0]
    decimal_str = parts[1] if len(parts) > 1 else ""
    
    def parse_section(s):
        val = 0; curr = 0; unit_val = 0
        for c in s:
            if c in CN_NUM: curr = CN_NUM[c]
            elif c in CN_UNIT:
                if c in ['万', '亿']: val = (val + unit_val + curr) * CN_UNIT[c]; unit_val = 0; curr = 0
                else: unit_val += curr * CN_UNIT[c]; curr = 0
        return val + unit_val + curr
    
    total = parse_section(integer_str)
    dec = 0.0; curr = 0
    for c in decimal_str:
        if c in CN_NUM: curr = CN_NUM[c]
        elif c == '角': dec += curr * 0.1; curr = 0
        elif c == '分': dec += curr * 0.01; curr = 0
    return round(total + dec, 2)

def find_amount_strict(text):
    """严格金额提取"""
    if not text: return 0.0, "空白"
    
    # 1. 大写优先
    up_m = re.search(r'(?:价税合计|大写|金额).*?([零壹贰叁肆伍陆柒捌玖拾佰仟万亿圆角分整]+)', text)
    amt_up = 0.0
    if up_m:
        try: amt_up = cn_upper_to_float(up_m.group(1))
        except: pass
    
    # 2. 小写校验
    lo_m = re.search(r'(?:小写|¥|￥|合计)[^0-9\.]*([0-9]{1,3}(?:,[0-9]{3})*\.[0-9]{2})', text)
    amt_lo = 0.0
    if lo_m:
        try: 
            v = float(lo_m.group(1).replace(",", ""))
            if 0.01 <= v <= 5000000: amt_lo = v
        except: pass

    # 3. 兜底 (全文最大值)
    if amt_up == 0 and amt_lo == 0:
        matches = re.findall(r'(\d{1,3}(?:,\d{3})*\.\d{2})', text)
        valid = []
        for m in matches:
            try:
                v = float(m.replace(",", ""))
                if 0.01 <= v <= 5000000 and v not in [0.06, 0.03, 0.13, 0.01, 1.00]: valid.append(v)
            except: continue
        if valid: amt_lo = max(valid)

    if amt_up > 0:
        if amt_lo > 0 and abs(amt_up - amt_lo) > 0.1:
            return amt_up, f"⚠️ 大小写不符({amt_up} vs {amt_lo})"
        return amt_up, "正常"
    
    if amt_lo > 0: return amt_lo, "使用小写"
    return 0.0, "警告:未读到金额"

def extract_seller_name_smart(text):
    suffix = r"[\u4e00-\u9fa5()（）]{2,30}(?:公司|事务所|酒店|旅行社|经营部|服务部|分行|支行|馆|店|处|中心)"
    candidates = list(set(re.findall(suffix, text)))
    blacklist = ["税务局", "财政部", "购买方", "开户行", "银行", "地址", "电话", "纳税人", "适用税率"]
    filtered = [c for c in candidates if not any(b in c for b in blacklist) and len(c) >= 4]
    return max(filtered, key=len) if filtered else ""

def parse_xml_invoice_data(xml_path):
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        def g(path): return root.find(path).text if root.find(path) is not None else ""
        num = g(".//TaxSupervisionInfo/InvoiceNumber") or g(".//InvoiceNumber") or g(".//Fphm")
        date = g(".//TaxSupervisionInfo/IssueTime") or g(".//IssueTime") or g(".//Kprq")
        seller = g(".//SellerInformation/SellerName") or g(".//Xfmc")
        amt_str = g(".//BasicInformation/TotalTax-includedAmount") or g(".//TotalTax-includedAmount") or g(".//TotalAmount")
        amount = float(amt_str.replace(',', '')) if amt_str else 0.0
        return {"num": num, "date": format_date(date), "seller": seller, "amount": amount}
    except: return None

def extract_data_from_pdf_simple(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as p:
            if not p.pages: return None
            raw = p.pages[0].extract_text()
            if not raw or len(raw.strip()) < 10: 
                return {"发票号码":"", "价税合计":0.0, "文件名":os.path.basename(pdf_path), "备注":"⚠️ 纯图/扫描件"}
            
            text = normalize_text(raw)
            num = ""
            m = re.search(r'(\d{20})', text)
            if m: num = m.group(1)
            else:
                m8 = re.search(r'(?:号码|No)[:|]?(\d{8,})', text)
                if m8: num = m8.group(1)
            
            date = ""
            md = re.search(r'(\d{4}[-年/.]\d{1,2}[-月/.]\d{1,2}日?)', text)
            if md: date = format_date(md.group(1))
            
            amt, note = find_amount_strict(text)
            seller = extract_seller_name_smart(text)
            
            if not num and amt > 0: note = "无发票号-" + note
            return {"发票号码": num, "开票日期": date, "销售方名称": seller, "价税合计": amt, "文件名": os.path.basename(pdf_path), "备注": note}
    except: return None

# ==========================================
# 4. 校验引擎 (Verifier)
# ==========================================

class InvoiceVerifier:
    def __init__(self, processed_df):
        self.processed_nums = {} 
        self.processed_attrs = {}
        for _, row in processed_df.iterrows():
            p_num = str(row.get('发票号码', '')).strip()
            p_amt = float(row.get('价税合计', 0))
            p_date = str(row.get('开票日期', '')).strip()
            
            if p_num and len(p_num) > 6:
                self.processed_nums[p_num] = {'amount': p_amt}
            
            attr_key = (f"{p_amt:.2f}", p_date)
            self.processed_attrs[attr_key] = True

    def check(self, raw_info):
        raw_num = str(raw_info.get('num') or raw_info.get('发票号码') or '').strip()
        raw_amt = float(raw_info.get('amount') or raw_info.get('价税合计') or 0)
        raw_date = str(raw_info.get('date') or raw_info.get('开票日期') or '').strip()
        
        # 1. 号码匹配
        if raw_num and len(raw_num) > 6:
            if raw_num in self.processed_nums: return True
        
        # 2. 金额+日期匹配
        if raw_amt > 0:
            if (f"{raw_amt:.2f}", raw_date) in self.processed_attrs: return True
        
        return False

# ==========================================
# 5. 主流程
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
    
    # 预扫描分类
    for pdf in pdf_files:
        try:
            with pdfplumber.open(pdf) as p:
                if not p.pages: continue
                text = normalize_text(p.pages[0].extract_text())
                amt, _ = find_amount_strict(text)
                folder = os.path.dirname(pdf)
                if is_trip_file(os.path.basename(pdf), text):
                    trip_pool.append({'path': pdf, 'amount': amt, 'folder': folder, 'used': False})
                else:
                    invoice_pdf_pool.append({'path': pdf, 'amount': amt, 'folder': folder})
        except: pass

    excel_rows = []
    idx = 1
    processed_source_files = set()

    # --- A. XML 处理 ---
    for xml in xml_files:
        info = parse_xml_invoice_data(xml)
        if not info: continue
        processed_source_files.add(os.path.abspath(xml))
        
        row = {"序号": idx, "发票号码": info['num'], "开票日期": info['date'],
               "销售方名称": info['seller'], "价税合计": info['amount'], 
               "数据来源": "XML", "文件名": os.path.basename(xml), "备注": "正常"}
        
        folder = os.path.dirname(xml)
        target_pdf = None
        cands = [p['path'] for p in invoice_pdf_pool if p['folder'] == folder]
        xml_base = os.path.splitext(os.path.basename(xml))[0]
        
        for p in cands:
            if xml_base in os.path.basename(p) or (info['num'] and info['num'] in os.path.basename(p)):
                target_pdf = p; break
        
        if target_pdf:
            processed_source_files.add(os.path.abspath(target_pdf))
            # 智能匹配
            matched_trip, match_remark = get_matching_trip_advanced(
                info['amount'], os.path.basename(target_pdf), folder, trip_pool
            )
            
            if matched_trip:
                matched_trip['used'] = True
                processed_source_files.add(os.path.abspath(matched_trip['path']))
                try:
                    merger = PdfWriter()
                    merger.append(target_pdf); merger.append(matched_trip['path'])
                    safe_name = f"{info['num']}_{info['amount']}.pdf".replace(':','').replace('/','_')
                    merger.write(os.path.join(merged_dir, safe_name)); merger.close()
                    row['备注'] = match_remark
                except:
                    shutil.copy2(target_pdf, os.path.join(noxml_dir, os.path.basename(target_pdf)))
                    row['备注'] = "合并失败-保留原件"
            else:
                shutil.copy2(target_pdf, os.path.join(noxml_dir, os.path.basename(target_pdf)))
        else:
             row['备注'] = "仅XML(缺PDF)"
        
        excel_rows.append(row); idx += 1

    # --- B. PDF 处理 ---
    for inv in invoice_pdf_pool:
        if os.path.abspath(inv['path']) in processed_source_files: continue
        
        data = extract_data_from_pdf_simple(inv['path'])
        if not data: continue
        processed_source_files.add(os.path.abspath(inv['path']))
        
        folder = inv['folder']
        matched_trip, match_remark = get_matching_trip_advanced(
            inv['amount'], os.path.basename(inv['path']), folder, trip_pool
        )
        
        if matched_trip:
            matched_trip['used'] = True
            processed_source_files.add(os.path.abspath(matched_trip['path']))
            try:
                merger = PdfWriter()
                merger.append(inv['path']); merger.append(matched_trip['path'])
                num = data.get('发票号码', 'NoNum')
                safe_name = f"{num}_{inv['amount']}.pdf".replace(':','').replace('/','_')
                merger.write(os.path.join(merged_dir, safe_name)); merger.close()
                data['备注'] = match_remark
            except:
                shutil.copy2(inv['path'], os.path.join(noxml_dir, os.path.basename(inv['path'])))
                data['备注'] = "合并失败-保留原件"
        else:
            shutil.copy2(inv['path'], os.path.join(noxml_dir, os.path.basename(inv['path'])))
            
        data['序号'] = idx; excel_rows.append(data); idx += 1

    # --- C. 剩余行程单 ---
    for t in trip_pool:
        if not t['used']:
            processed_source_files.add(os.path.abspath(t['path']))
            try: shutil.copy2(t['path'], os.path.join(noxml_dir, os.path.basename(t['path'])))
            except: pass

    # --- D. 生成与核对 ---
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
        df_disp = pd.concat([df_result, pd.DataFrame([sum_row])], ignore_index=True)
        excel_path = os.path.join(output_dir, 'Summary_Final.xlsx')
        df_disp.to_excel(excel_path, index=False)

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
                if "纯图" in raw_info.get('备注', ''): is_missing = True
                elif verifier.check(raw_info): is_missing = False
            
            if is_missing: missing_files.append(f)

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
            rows = []
            for n in z.namelist():
                if n.endswith('.pdf'):
                    m = re.match(r'(\d+)_([\d\.]+)\.pdf', os.path.basename(n))
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
            
            if raw_info and verifier.check(raw_info): matched_count += 1
            else: missing.append(fp)
    
    zip_p = None
    if missing:
        zip_p = os.path.join(out_dir, "Manual_Missing.zip")
        with zipfile.ZipFile(zip_p, 'w', zipfile.ZIP_DEFLATED) as z:
            for m in missing: z.write(m, os.path.basename(m))
            
    return matched_count, len(missing), zip_p

# ==========================================
# 7. Streamlit UI
# ==========================================

def main():
    st.set_page_config(page_title="发票无忧 V14 (完美匹配版)", layout="wide")
    st.title("🧾 发票无忧 V14 (文件名优先 + 人工复核)")

    tab1, tab2 = st.tabs(["🚀 一键处理", "🔍 手动复核"])

    with tab1:
        st.info("策略：1.文件名核心匹配(优先) 2.金额匹配 3.自动高亮金额不符项")
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
        st.write("反向核对工具")
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