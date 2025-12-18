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
    if not text: return ""
    return text.replace(" ", "").replace("\n", "").replace("\r", "")\
               .replace("：", ":").replace("￥", "¥")\
               .replace("（", "(").replace("）", ")")

def find_max_valid_amount(text):
    """提取金额"""
    matches = re.findall(r'(\d{1,3}(?:,\d{3})*\.\d{2})', text)
    valid_amounts = []
    for m in matches:
        try:
            val = float(m.replace(",", ""))
            # 排除常见干扰：税率(0.06/0.13)、数量(1.00)、过小的金额
            if 0.01 <= val <= 5000000 and val not in [0.06, 0.03, 0.13, 0.01, 1.00]:
                valid_amounts.append(val)
        except: continue
    return max(valid_amounts) if valid_amounts else 0.0

def extract_seller_name_smart(text):
    """提取销售方"""
    suffix_pattern = r"[\u4e00-\u9fa5()（）]{2,30}(?:公司|事务所|酒店|旅行社|经营部|服务部|分行|支行|馆|店|处|中心)"
    candidates = list(set(re.findall(suffix_pattern, text)))
    blacklist = ["税务局", "财政部", "购买方", "开户行", "银行", "地址", "电话", "统一社会信用", "纳税人", "适用税率", "密码区"]
    filtered = [c for c in candidates if not any(b in c for b in blacklist) and len(c) >= 4]
    return max(filtered, key=len) if filtered else ""

def is_trip_file(filename, text=None):
    """判断是否为行程单"""
    fn = filename.lower()
    if "行程" in fn or "trip" in fn or "报销" in fn:
        if text:
            clean = normalize_text(text)
            if "发票代码" in clean or "发票号码" in clean or "电子发票" in clean:
                return False
        return True
    return False

# ==========================================
# 2. 解析函数
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
        # 修复金额逗号问题
        amount = float(amt_str.replace(',', '')) if amt_str else 0.0

        return {"num": num, "date": date, "seller": seller, "amount": amount}
    except: return None

def extract_data_from_pdf_simple(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as p:
            if not p.pages: return None
            raw = p.pages[0].extract_text()
            if not raw: return None
            text = normalize_text(raw)
            
            num = ""
            m20 = re.search(r'(\d{20})', text)
            if m20: num = m20.group(1)
            else:
                m8 = re.search(r'(?:号码|No)[:|]?(\d{8,})', text)
                if m8: num = m8.group(1)
            
            date = ""
            md = re.search(r'(\d{4}[-年/.]\d{1,2}[-月/.]\d{1,2}日?)', text)
            if md: date = md.group(1)
            
            amt = find_max_valid_amount(text)
            seller = extract_seller_name_smart(text)
            
            return {
                "发票号码": num, "开票日期": date, "销售方名称": seller,
                "价税合计": amt, "数据来源": "PDF识别", "文件名": os.path.basename(pdf_path),
                "备注": "正常" if amt > 0 else "警告:未读到金额"
            }
    except: return None

# ==========================================
# 3. 核心处理流程 (含自动核对)
# ==========================================

def run_process_pipeline(input_root_dir, output_dir):
    """
    执行处理并在最后进行自动核对
    返回: excel路径, 合并文件夹, 独立发票文件夹, 遗漏文件列表
    """
    merged_pdf_dir = os.path.join(output_dir, 'Merged_PDFs')
    no_xml_pdf_dir = os.path.join(output_dir, 'No_XML_PDFs')
    os.makedirs(merged_pdf_dir, exist_ok=True)
    os.makedirs(no_xml_pdf_dir, exist_ok=True)

    # 1. 扫描所有文件
    all_files = []
    for root, dirs, files in os.walk(input_root_dir):
        for f in files: all_files.append(os.path.join(root, f))
    
    xml_files = [f for f in all_files if f.lower().endswith('.xml')]
    pdf_files = [f for f in all_files if f.lower().endswith('.pdf')]
    
    # 2. 建立行程单池 & 发票候选池 (区分 Scope)
    trip_pool = []
    invoice_pdf_pool = []
    
    for pdf in pdf_files:
        try:
            with pdfplumber.open(pdf) as p:
                if not p.pages: continue
                text = normalize_text(p.pages[0].extract_text())
                amt = find_max_valid_amount(text)
                folder = os.path.dirname(pdf)
                
                if is_trip_file(os.path.basename(pdf), text):
                    trip_pool.append({'path': pdf, 'amount': amt, 'folder': folder, 'used': False})
                else:
                    invoice_pdf_pool.append({'path': pdf, 'amount': amt, 'folder': folder})
        except: pass

    excel_rows = []
    idx = 1
    
    # 【核对关键】记录哪些原始文件被成功使用了
    processed_source_files = set()

    # --- 阶段 A: XML 发票 ---
    for xml in xml_files:
        info = parse_xml_invoice_data(xml)
        if not info: continue
        
        # 标记 XML 为已处理
        processed_source_files.add(os.path.abspath(xml))
        
        row = {
            "序号": idx, "发票号码": info['num'], "开票日期": info['date'],
            "销售方名称": info['seller'], "价税合计": info['amount'], 
            "数据来源": "XML", "文件名": os.path.basename(xml), "备注": "正常"
        }
        
        # 找 PDF
        folder = os.path.dirname(xml)
        target_pdf = None
        
        # Scope 匹配
        cands = [p['path'] for p in invoice_pdf_pool if p['folder'] == folder]
        xml_base = os.path.splitext(os.path.basename(xml))[0]
        
        for p in cands:
            if xml_base in os.path.basename(p) or (info['num'] and info['num'] in os.path.basename(p)):
                target_pdf = p
                break
        
        if target_pdf:
            processed_source_files.add(os.path.abspath(target_pdf))
            
            # 找行程单
            matched_trip = None
            trips = [t for t in trip_pool if t['folder'] == folder and not t['used']]
            for t in trips:
                if abs(t['amount'] - info['amount']) < 0.05:
                    matched_trip = t
                    t['used'] = True
                    break
            
            if matched_trip:
                processed_source_files.add(os.path.abspath(matched_trip['path']))
                try:
                    merger = PdfWriter()
                    merger.append(target_pdf)
                    merger.append(matched_trip['path'])
                    safe_name = f"{info['num']}_{info['amount']}.pdf".replace(':','').replace('/','_')
                    merger.write(os.path.join(merged_pdf_dir, safe_name))
                    merger.close()
                    row['备注'] = "已合并行程单"
                except:
                    # 合并失败，复制原件作为兜底
                    shutil.copy2(target_pdf, os.path.join(no_xml_pdf_dir, os.path.basename(target_pdf)))
                    row['备注'] = "合并失败-保留原件"
            else:
                shutil.copy2(target_pdf, os.path.join(no_xml_pdf_dir, os.path.basename(target_pdf)))
        
        excel_rows.append(row)
        idx += 1

    # --- 阶段 B: 无 XML 的 PDF ---
    for inv in invoice_pdf_pool:
        if os.path.abspath(inv['path']) in processed_source_files: continue
        
        data = extract_data_from_pdf_simple(inv['path'])
        if not data: continue
        
        processed_source_files.add(os.path.abspath(inv['path']))
        
        matched_trip = None
        folder = inv['folder']
        trips = [t for t in trip_pool if t['folder'] == folder and not t['used']]
        for t in trips:
            if inv['amount'] > 0 and abs(t['amount'] - inv['amount']) < 0.05:
                matched_trip = t
                t['used'] = True
                break
        
        if matched_trip:
            processed_source_files.add(os.path.abspath(matched_trip['path']))
            try:
                merger = PdfWriter()
                merger.append(inv['path'])
                merger.append(matched_trip['path'])
                num = data.get('发票号码', 'NoNum')
                safe_name = f"{num}_{inv['amount']}.pdf".replace(':','').replace('/','_')
                merger.write(os.path.join(merged_pdf_dir, safe_name))
                merger.close()
                data['备注'] = "已合并行程单"
                if data['价税合计'] == 0: data['价税合计'] = inv['amount']
            except:
                shutil.copy2(inv['path'], os.path.join(no_xml_pdf_dir, os.path.basename(inv['path'])))
                data['备注'] = "合并失败-保留原件"
        else:
            shutil.copy2(inv['path'], os.path.join(no_xml_pdf_dir, os.path.basename(inv['path'])))
        
        data['序号'] = idx
        excel_rows.append(data)
        idx += 1

    # --- 阶段 C: 剩余行程单 ---
    for t in trip_pool:
        if not t['used']:
            processed_source_files.add(os.path.abspath(t['path']))
            try: shutil.copy2(t['path'], os.path.join(no_xml_pdf_dir, os.path.basename(t['path'])))
            except: pass

    # --- 阶段 D: 自动核对 (找出遗漏文件) ---
    missing_files = []
    # 过滤只检查 pdf 和 xml
    check_exts = ('.pdf', '.xml')
    for f in all_files:
        if f.lower().endswith(check_exts):
            if os.path.abspath(f) not in processed_source_files:
                missing_files.append(f)

    # 生成 Excel
    excel_path = None
    if excel_rows:
        df = pd.DataFrame(excel_rows)
        cols = ["序号", "发票号码", "开票日期", "销售方名称", "价税合计", "数据来源", "备注", "文件名"]
        for c in cols: 
            if c not in df.columns: df[c] = ""
        df = df[cols]
        df['价税合计'] = pd.to_numeric(df['价税合计'], errors='coerce').fillna(0.0)
        sum_row = {"序号": "总计", "价税合计": df['价税合计'].sum(), "销售方名称": f"共 {len(df)} 张"}
        df = pd.concat([df, pd.DataFrame([sum_row])], ignore_index=True)
        excel_path = os.path.join(output_dir, 'Summary_Final.xlsx')
        df.to_excel(excel_path, index=False)

    return excel_path, merged_pdf_dir, no_xml_pdf_dir, missing_files

# ==========================================
# 4. 手动核对功能 (Tab 2)
# ==========================================
def run_manual_check(raw_dir, proc_zip_path, out_dir):
    """基于文件名的手动核对"""
    # 1. 读取已处理发票号
    processed_nums = set()
    with zipfile.ZipFile(proc_zip_path, 'r') as z:
        for n in z.namelist():
            base = os.path.basename(n)
            m = re.search(r'^(\d{8,})', base)
            if m: processed_nums.add(m.group(1))

    # 2. 扫描原始文件
    missing = []
    matched_count = 0
    
    for root, _, files in os.walk(raw_dir):
        for f in files:
            if not f.lower().endswith(('.pdf', '.xml')): continue
            fp = os.path.join(root, f)
            
            # 简易提取发票号
            num = None
            try:
                if f.endswith('.xml'):
                    info = parse_xml_invoice_data(fp)
                    if info: num = info['num']
                else:
                    data = extract_data_from_pdf_simple(fp)
                    if data: num = data['发票号码']
            except: pass
            
            # 判断
            # 如果是行程单(Trip)，且没有被合并(不在zip里体现)，可能无法直接通过文件名判断
            # 这里主要核对主发票
            if num and num in processed_nums:
                matched_count += 1
            else:
                # 只有当它是发票且没找到时才算Missing
                # 或是行程单且没找到
                missing.append(fp)
    
    # 打包
    zip_p = None
    if missing:
        zip_p = os.path.join(out_dir, "Manual_Missing.zip")
        with zipfile.ZipFile(zip_p, 'w', zipfile.ZIP_DEFLATED) as z:
            for m in missing: z.write(m, os.path.basename(m))
            
    return matched_count, len(missing), zip_p

# ==========================================
# 5. Streamlit 主界面
# ==========================================

def main():
    st.set_page_config(page_title="发票无忧 V10 (终极版)", layout="wide")
    st.title("🧾 发票无忧 V10 (含自动核对与遗漏打包)")

    tab1, tab2 = st.tabs(["🚀 一键处理 (自动核对)", "🔍 手动复核 (旧包审计)"])

    # --- Tab 1: 自动处理 + 自动核对 ---
    with tab1:
        st.info("上传 ZIP/文件夹，系统会自动：1.隔离作用域匹配行程单 2.生成汇总 3.自动找出遗漏文件并打包")
        uploaded_files = st.file_uploader("上传文件", type=['zip', 'xml', 'pdf'], accept_multiple_files=True, key="u1")

        if uploaded_files and st.button("开始处理", key="b1"):
            with st.spinner('正在全流程处理...'):
                with tempfile.TemporaryDirectory() as temp_dir:
                    input_root = os.path.join(temp_dir, "input")
                    os.makedirs(input_root, exist_ok=True)
                    
                    # 1. 物理隔离保存
                    for i, up in enumerate(uploaded_files):
                        scope_dir = os.path.join(input_root, f"scope_{i}")
                        os.makedirs(scope_dir, exist_ok=True)
                        save_path = os.path.join(scope_dir, up.name)
                        with open(save_path, "wb") as f: f.write(up.getbuffer())
                        if up.name.endswith('.zip'):
                            extract_zip_with_encoding(save_path, scope_dir)
                            os.remove(save_path)
                    
                    # 2. 运行管道
                    out_dir = os.path.join(temp_dir, "output")
                    excel, merged, noxml, missing_list = run_process_pipeline(input_root, out_dir)
                    
                    # 3. 结果展示
                    st.success("✅ 处理完成！")
                    
                    # 统计
                    col1, col2 = st.columns(2)
                    if excel:
                        df = pd.read_excel(excel)
                        count = len(df) - 1 # 减去总计行
                        col1.metric("成功匹配发票", f"{count} 张")
                        st.dataframe(df.tail(3))
                    
                    # 遗漏处理
                    col2.metric("遗漏文件", f"{len(missing_list)} 个", delta_color="inverse")
                    if missing_list:
                        st.warning("⚠️ 检测到有文件未被处理（可能是损坏、加密或非发票文件），已打包如下：")
                        missing_zip = os.path.join(temp_dir, "Missing_Files.zip")
                        with zipfile.ZipFile(missing_zip, 'w', zipfile.ZIP_DEFLATED) as z:
                            for mf in missing_list:
                                z.write(mf, f"遗漏文件/{os.path.basename(mf)}")
                        with open(missing_zip, "rb") as f:
                            st.download_button("📥 下载遗漏文件包 (Missing.zip)", f, "Missing_Files.zip", type="primary")

                    # 主结果打包
                    if excel:
                        res_zip = os.path.join(temp_dir, "Result.zip")
                        with zipfile.ZipFile(res_zip, 'w', zipfile.ZIP_DEFLATED) as z:
                            z.write(excel, "汇总表.xlsx")
                            for r, _, fs in os.walk(merged):
                                for f in fs: z.write(os.path.join(r, f), f"合并后发票/{f}")
                            for r, _, fs in os.walk(noxml):
                                for f in fs: z.write(os.path.join(r, f), f"独立发票/{f}")
                        
                        with open(res_zip, "rb") as f:
                            st.download_button("📥 下载处理结果 (Result.zip)", f, "Invoices_Result.zip")

    # --- Tab 2: 手动核对 ---
    with tab2:
        st.write("用于核对**以前处理过的**结果包。")
        c1, c2 = st.columns(2)
        raw_ups = c1.file_uploader("1. 上传原始发票 (ZIP/PDF)", type=['zip','pdf'], accept_multiple_files=True, key="u2")
        proc_zip = c2.file_uploader("2. 上传已处理 Result.zip", type=['zip'], key="u3")
        
        if raw_ups and proc_zip and st.button("开始核对", key="b2"):
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
                st.metric("匹配成功", match)
                st.metric("遗漏/未匹配", miss)
                if mzip:
                    with open(mzip, "rb") as f:
                        st.download_button("下载未匹配文件", f, "Unmatched.zip")

if __name__ == "__main__":
    main()