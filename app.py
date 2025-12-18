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
# 1. 基础工具函数 (解压、清洗、通用提取)
# ==========================================

def extract_zip_with_encoding(zip_path, extract_to):
    """解压 ZIP 并修复中文乱码 (CP437 -> GBK 自动转换)"""
    with zipfile.ZipFile(zip_path, 'r') as z:
        for file_info in z.infolist():
            try:
                # 尝试修复文件名编码
                if file_info.flag_bits & 0x800 == 0:
                    original_name = file_info.filename.encode('cp437').decode('gbk')
                else:
                    original_name = file_info.filename
            except:
                try: original_name = file_info.filename.encode('utf-8').decode('utf-8')
                except: original_name = file_info.filename

            # 过滤掉 macOS 隐藏文件
            if "__MACOSX" in original_name or ".DS_Store" in original_name:
                continue

            target_path = os.path.join(extract_to, original_name)
            
            # 防止路径穿越，确保父目录存在
            parent_dir = os.path.dirname(target_path)
            if parent_dir and not os.path.exists(parent_dir):
                os.makedirs(parent_dir, exist_ok=True)
                
            # 只解压文件，不解压纯文件夹条目
            if not original_name.endswith('/'):
                with z.open(file_info) as source, open(target_path, "wb") as target:
                    shutil.copyfileobj(source, target)

def normalize_text(text):
    """清洗文本：去空格、换行、全角转半角"""
    if not text: return ""
    return text.replace(" ", "").replace("\n", "").replace("\r", "")\
               .replace("：", ":").replace("￥", "¥")\
               .replace("（", "(").replace("）", ")")

def find_max_valid_amount(text):
    """
    从文本中提取所有看起来像金额的数字，取最大值作为价税合计。
    排除日期、税率、数量等干扰。
    """
    # 匹配 123.45 或 1,234.56 格式
    matches = re.findall(r'(\d{1,3}(?:,\d{3})*\.\d{2})', text)
    valid_amounts = []
    for m in matches:
        try:
            val = float(m.replace(",", ""))
            # 过滤逻辑：金额通常在 0.01 到 500万之间
            # 排除常见的税率 0.06, 0.03, 0.13, 0.01 和数量 1.00
            if 0.01 <= val <= 5000000 and val not in [0.06, 0.03, 0.13, 0.01, 1.00]:
                valid_amounts.append(val)
        except: continue
    
    return max(valid_amounts) if valid_amounts else 0.0

def extract_seller_name_smart(text):
    """智能提取销售方名称"""
    # 匹配以特定后缀结尾的中文名称
    suffix_pattern = r"[\u4e00-\u9fa5()（）]{2,30}(?:公司|事务所|酒店|旅行社|经营部|服务部|分行|支行|馆|店|处|中心)"
    candidates = list(set(re.findall(suffix_pattern, text)))
    
    # 黑名单过滤
    blacklist = ["税务局", "财政部", "购买方", "开户行", "银行", "地址", "电话", "统一社会信用", "纳税人", "适用税率"]
    filtered = [c for c in candidates if not any(b in c for b in blacklist) and len(c) >= 4]
    
    if not filtered: return ""
    # 通常取最长的名字作为销售方全称
    return max(filtered, key=len)

def is_trip_file(filename, text=None):
    """判断是否为行程单/报销单"""
    fn = filename.lower()
    # 特征 1: 文件名包含关键字
    if "行程" in fn or "trip" in fn or "报销" in fn:
        # 特征 2: 内容排除发票特征 (防止文件名叫行程单但其实是发票)
        if text:
            clean_text = normalize_text(text)
            # 如果内容里有明确的"发票代码"、"发票号码"、"电子发票"，则即使文件名有行程也视为发票
            if "发票代码" in clean_text or "发票号码" in clean_text or "电子发票" in clean_text:
                return False
        return True
    return False

# ==========================================
# 2. 核心解析函数 (XML & PDF) - 已补全
# ==========================================

def parse_xml_invoice_data(xml_path):
    """
    完整解析 XML 数据
    适配：数电票（税务局）、航信、百望云等不同结构
    """
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        
        # 辅助函数：安全获取节点文本
        def g(path):
            node = root.find(path)
            return node.text if node is not None else ""

        # 1. 提取发票号码 (不同规范路径不同)
        num = g(".//TaxSupervisionInfo/InvoiceNumber")
        if not num: num = g(".//InvoiceNumber")
        if not num: num = g(".//Fphm") # 部分旧接口
        
        # 2. 提取日期
        date = g(".//TaxSupervisionInfo/IssueTime")
        if not date: date = g(".//IssueTime")
        if not date: date = g(".//Kprq")
        
        # 3. 提取销售方
        seller = g(".//SellerInformation/SellerName")
        if not seller: seller = g(".//Xfmc")
        
        # 4. 提取金额 (价税合计)
        amt_str = g(".//BasicInformation/TotalTax-includedAmount")
        if not amt_str: amt_str = g(".//TotalTax-includedAmount")
        if not amt_str: amt_str = g(".//TotalAmount") # 兼容
        if not amt_str: amt_str = g(".//Jshj")
        
        amount = float(amt_str) if amt_str else 0.0

        return {
            "num": num,
            "date": date,
            "seller": seller,
            "amount": amount
        }
    except Exception as e:
        # print(f"XML Parse Error {xml_path}: {e}")
        return None

def extract_data_from_pdf_simple(pdf_path):
    """
    从 PDF 中提取基础发票数据
    """
    try:
        with pdfplumber.open(pdf_path) as p:
            if not p.pages: return None
            # 获取第一页文本
            raw_text = p.pages[0].extract_text()
            if not raw_text: return None
            
            clean_text = normalize_text(raw_text)
            
            # 1. 提取发票号码 (优先找20位全电号码，其次找普通发票号)
            num = ""
            m_20 = re.search(r'(\d{20})', clean_text)
            if m_20:
                num = m_20.group(1)
            else:
                m_8 = re.search(r'(?:号码|No)[:|]?(\d{8,})', clean_text)
                if m_8: num = m_8.group(1)
            
            # 2. 提取日期
            date = ""
            m_date = re.search(r'(\d{4}年\d{1,2}月\d{1,2}日)', clean_text)
            if m_date: date = m_date.group(1)
            
            # 3. 提取金额
            amount = find_max_valid_amount(clean_text)
            
            # 4. 提取销售方
            seller = extract_seller_name_smart(clean_text)
            
            return {
                "发票号码": num,
                "开票日期": date,
                "销售方名称": seller,
                "价税合计": amount,
                "数据来源": "PDF识别",
                "文件名": os.path.basename(pdf_path),
                "备注": "正常" if amount > 0 else "警告:未读到金额"
            }
    except Exception:
        return None

# ==========================================
# 3. 业务流程逻辑 (严格闭环匹配)
# ==========================================

def run_process_pipeline(input_root_dir, output_dir):
    """
    input_root_dir: 包含多个独立 scope 文件夹的根目录
    output_dir: 结果输出目录
    """
    merged_pdf_dir = os.path.join(output_dir, 'Merged_PDFs')
    no_xml_pdf_dir = os.path.join(output_dir, 'No_XML_PDFs')
    os.makedirs(merged_pdf_dir, exist_ok=True)
    os.makedirs(no_xml_pdf_dir, exist_ok=True)

    # 1. 遍历所有文件，建立索引
    # 注意：这里会遍历 input_root_dir 下的所有子文件夹 (Scope)
    all_files = []
    for root, dirs, files in os.walk(input_root_dir):
        for f in files:
            all_files.append(os.path.join(root, f))
    
    xml_files = [f for f in all_files if f.lower().endswith('.xml')]
    pdf_files = [f for f in all_files if f.lower().endswith('.pdf')]
    
    # 2. 建立 Trip Pool (行程单池)
    # 关键结构: {'path': abs_path, 'amount': 123.45, 'folder': dir_path, 'used': False}
    trip_pool = []
    invoice_pdf_pool = [] # 没有XML的发票PDF候选
    
    # 预扫描 PDF
    for pdf in pdf_files:
        try:
            with pdfplumber.open(pdf) as p:
                if not p.pages: continue
                text = normalize_text(p.pages[0].extract_text())
                amount = find_max_valid_amount(text)
                folder = os.path.dirname(pdf) # 获取物理文件夹路径 (Scope)
                
                if is_trip_file(os.path.basename(pdf), text):
                    trip_pool.append({'path': pdf, 'amount': amount, 'folder': folder, 'used': False})
                else:
                    # 只要不是行程单，都视为发票候选
                    invoice_pdf_pool.append({'path': pdf, 'amount': amount, 'folder': folder})
        except: pass

    excel_rows = []
    idx = 1
    processed_invoice_pdfs = set() # 记录已被处理的PDF路径

    # --- 阶段 A: 优先处理 XML (准确度最高) ---
    for xml in xml_files:
        inv_info = parse_xml_invoice_data(xml)
        if not inv_info: continue
        
        row = {
            "序号": idx, "发票号码": inv_info['num'], "开票日期": inv_info['date'],
            "销售方名称": inv_info['seller'], "价税合计": inv_info['amount'], 
            "数据来源": "XML", "文件名": os.path.basename(xml), "备注": "正常"
        }
        
        # 1. 在同目录下找对应的发票 PDF
        xml_folder = os.path.dirname(xml) # 锁定 Scope
        target_invoice_pdf = None
        
        # 筛选：Scope 必须相同
        potential_invs = [p['path'] for p in invoice_pdf_pool if p['folder'] == xml_folder]
        xml_base = os.path.splitext(os.path.basename(xml))[0]
        
        for p_path in potential_invs:
            p_name = os.path.basename(p_path)
            # 匹配逻辑：同名 OR 包含发票号
            if xml_base in p_name or (inv_info['num'] and inv_info['num'] in p_name):
                target_invoice_pdf = p_path
                break
        
        # 如果找到了 PDF
        if target_invoice_pdf:
            processed_invoice_pdfs.add(target_invoice_pdf)
            
            # 2. 在同目录下找匹配的行程单 (Strict Match)
            matched_trip = None
            candidate_trips = [t for t in trip_pool if t['folder'] == xml_folder and not t['used']]
            
            for trip in candidate_trips:
                # 匹配逻辑：金额一致 (误差 < 0.05)
                if abs(trip['amount'] - inv_info['amount']) < 0.05:
                    matched_trip = trip
                    trip['used'] = True
                    break
            
            if matched_trip:
                try:
                    merger = PdfWriter()
                    merger.append(target_invoice_pdf)
                    merger.append(matched_trip['path'])
                    # 安全文件名
                    safe_name = f"{inv_info['num']}_{inv_info['amount']}.pdf".replace(':','').replace('/','_')
                    merger.write(os.path.join(merged_pdf_dir, safe_name))
                    merger.close()
                    row['备注'] = "已合并行程单"
                except: pass
            else:
                # 没找到行程单，复制原 PDF 到 No_XML (作为未合并发票)
                try: shutil.copy2(target_invoice_pdf, os.path.join(no_xml_pdf_dir, os.path.basename(target_invoice_pdf)))
                except: pass
        
        excel_rows.append(row)
        idx += 1

    # --- 阶段 B: 处理无 XML 的 PDF 发票 ---
    for inv_pdf in invoice_pdf_pool:
        if inv_pdf['path'] in processed_invoice_pdfs: continue
        
        # 提取发票数据
        pdf_data = extract_data_from_pdf_simple(inv_pdf['path'])
        if not pdf_data: continue
        
        # 在同目录下找匹配行程单
        matched_trip = None
        folder = inv_pdf['folder'] # 锁定 Scope
        candidate_trips = [t for t in trip_pool if t['folder'] == folder and not t['used']]
        
        for trip in candidate_trips:
            # 金额必须有效(>0)且一致
            if inv_pdf['amount'] > 0 and abs(trip['amount'] - inv_pdf['amount']) < 0.05:
                matched_trip = trip
                trip['used'] = True
                break
        
        # 有匹配则合并，无匹配则保留原件
        if matched_trip:
            try:
                merger = PdfWriter()
                merger.append(inv_pdf['path'])
                merger.append(matched_trip['path'])
                
                num = pdf_data.get('发票号码', 'NoNum')
                amt = inv_pdf['amount']
                safe_name = f"{num}_{amt}.pdf".replace(':','').replace('/','_')
                merger.write(os.path.join(merged_pdf_dir, safe_name))
                merger.close()
                
                pdf_data['备注'] = "已合并行程单(PDF匹配)"
                # 确保 Excel 里金额是准确的（优先信 PDF 提取的，或者行程单的）
                if pdf_data['价税合计'] == 0: pdf_data['价税合计'] = amt
            except: pass
        else:
            try:
                shutil.copy2(inv_pdf['path'], os.path.join(no_xml_pdf_dir, os.path.basename(inv_pdf['path'])))
            except: pass
        
        # 补全序号并添加
        pdf_data['序号'] = idx
        excel_rows.append(pdf_data)
        idx += 1

    # --- 阶段 C: 兜底 (保留未使用的行程单) ---
    for trip in trip_pool:
        if not trip['used']:
            try: shutil.copy2(trip['path'], os.path.join(no_xml_pdf_dir, os.path.basename(trip['path'])))
            except: pass

    # 生成 Excel
    if excel_rows:
        df = pd.DataFrame(excel_rows)
        # 确保列顺序
        cols = ["序号", "发票号码", "开票日期", "销售方名称", "价税合计", "数据来源", "备注", "文件名"]
        for c in cols: 
            if c not in df.columns: df[c] = ""
        df = df[cols]
        
        # 格式化金额
        df['价税合计'] = pd.to_numeric(df['价税合计'], errors='coerce').fillna(0.0)
        
        # 添加总计行
        sum_row = {"序号": "总计", "价税合计": df['价税合计'].sum(), "销售方名称": f"共 {len(df)} 张"}
        df = pd.concat([df, pd.DataFrame([sum_row])], ignore_index=True)
        
        excel_path = os.path.join(output_dir, 'Summary_Final.xlsx')
        df.to_excel(excel_path, index=False)
        return excel_path, merged_pdf_dir, no_xml_pdf_dir
    
    return None, None, None

# ==========================================
# 4. Streamlit 界面
# ==========================================

def main():
    st.set_page_config(page_title="发票无忧 V8.0 (闭环版)", layout="wide")
    st.title("🧾 发票无忧 V8.0 (严格闭环匹配)")
    st.info("功能：上传 ZIP/文件夹，系统会自动在【同一个包内】严格匹配发票和行程单。")

    uploaded_files = st.file_uploader(
        "请上传发票 ZIP (支持多包上传)", 
        type=['zip', 'xml', 'pdf'], 
        accept_multiple_files=True
    )

    if uploaded_files and st.button("开始处理"):
        with st.spinner('正在分析 (保持文件包隔离)...'):
            with tempfile.TemporaryDirectory() as temp_dir:
                # input_root 是所有 scope 文件夹的父级
                input_root = os.path.join(temp_dir, "input_root")
                os.makedirs(input_root, exist_ok=True)
                
                # === 关键：为每个上传项建立独立文件夹 (Scope) ===
                for i, up_file in enumerate(uploaded_files):
                    # 文件夹名: index_filename
                    safe_foldername = f"scope_{i}_{re.sub(r'[^a-zA-Z0-9]', '_', up_file.name)}"
                    file_scope_dir = os.path.join(input_root, safe_foldername)
                    os.makedirs(file_scope_dir, exist_ok=True)
                    
                    save_path = os.path.join(file_scope_dir, up_file.name)
                    with open(save_path, "wb") as f:
                        f.write(up_file.getbuffer())
                    
                    # 如果是 ZIP，解压到当前 Scope
                    if up_file.name.lower().endswith('.zip'):
                        extract_zip_with_encoding(save_path, file_scope_dir)
                        os.remove(save_path) # 删除原 ZIP
                
                # 执行处理
                output_dir = os.path.join(temp_dir, "output")
                excel, merged, noxml = run_process_pipeline(input_root, output_dir)
                
                if excel:
                    st.success("处理完成！")
                    st.dataframe(pd.read_excel(excel).tail(5))
                    
                    # 打包结果
                    res_zip = os.path.join(temp_dir, "Result.zip")
                    with zipfile.ZipFile(res_zip, 'w', zipfile.ZIP_DEFLATED) as z:
                        z.write(excel, "汇总表.xlsx")
                        for r, _, fs in os.walk(merged):
                            for f in fs: z.write(os.path.join(r, f), f"合并后发票/{f}")
                        for r, _, fs in os.walk(noxml):
                            for f in fs: z.write(os.path.join(r, f), f"独立发票/{f}")
                            
                    with open(res_zip, "rb") as f:
                        st.download_button("下载结果包", f, "Invoices_Scoped.zip")
                else:
                    st.error("未找到有效发票。")

if __name__ == "__main__":
    main()