import streamlit as st
import pandas as pd
import io
import copy
import os
from datetime import datetime
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import Pt
from docx.oxml.ns import qn

# ==========================================
# 核心算法逻辑 (严格保持不变，修复了可能的空值处理)
# ==========================================
def format_date(value):
    if pd.isna(value): return ""
    if isinstance(value, (pd.Timestamp, datetime)):
        return f"{value.year}年{value.month}月{value.day}日"
    return str(value)

def set_dual_font_10pt(run):
    run.font.name = 'Arial'
    run.font.size = Pt(10)
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), '宋体')
    rFonts.set(qn('w:ascii'), 'Arial')
    rFonts.set(qn('w:hAnsi'), 'Arial')

def replace_text_strict_format(paragraph, replacements):
    if not paragraph.runs: return
    full_text = "".join(run.text for run in paragraph.runs)
    new_text = full_text
    for k, v in replacements.items():
        new_text = new_text.replace(str(k), str(v)) # 确保 k 也是字符串
    if new_text != full_text:
        paragraph.runs[0].text = new_text
        for i in range(1, len(paragraph.runs)):
            paragraph.runs[i].text = ""

def replace_all_content_keep_style(doc, replacements):
    for p in doc.paragraphs: replace_text_strict_format(p, replacements)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs: replace_text_strict_format(p, replacements)
    for node in doc.element.xpath("//w:t"):
        if node.text:
            text = node.text
            for k, v in replacements.items(): text = text.replace(str(k), str(v))
            node.text = text

def replace_header_keep_format(doc, replacements):
    def replace_in_header(header):
        for paragraph in header.paragraphs:
            for run in paragraph.runs:
                if not run.text: continue
                text = run.text
                for k, v in replacements.items():
                    if str(k) in text: text = text.replace(str(k), str(v))
                run.text = text
    for i, section in enumerate(doc.sections):
        try:
            section.header.is_linked_to_previous = False
            replace_in_header(section.header)
            # 兼容旧版本 docx 可能不存在的属性
            if hasattr(section, 'first_page_header'): replace_in_header(section.first_page_header)
            if hasattr(section, 'even_page_header'): replace_in_header(section.even_page_header)
        except: pass

def fill_specific_table(table, records):
    if not records or len(table.rows) < 2: return
    template_row = table.rows[1]
    # 清理除标题和模板行以外的所有行
    while len(table.rows) > 2: table._tbl.remove(table.rows[-1]._tr)
    for idx, record in enumerate(records, start=1):
        new_row_xml = copy.deepcopy(template_row._tr)
        table._tbl.append(new_row_xml)
        row = table.rows[-1]
        # 对应：序号, 文献信息, 检索说明, 排除理由, 备注
        row_data = [str(idx), record[1], record[2], record[3], record[4]]
        for i, val in enumerate(row_data):
            if i < len(row.cells):
                cell = row.cells[i]
                if not cell.paragraphs: cell.add_paragraph()
                p = cell.paragraphs[0]
                p.text = ""
                run = p.add_run(str(val) if not pd.isna(val) else "")
                set_dual_font_10pt(run)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    # 最后移除用于复制的模板行
    table._tbl.remove(template_row._tr)

# ==========================================
# 网站 UI 增强部分 (排版优化)
# ==========================================

st.set_page_config(page_title="ReportGen | PRER报告自动生成工具-CER中心", layout="wide")

# 注入 CSS 美化界面
st.markdown("""
    <style>
    .stAlert { border-radius: 10px; }
    .stButton>button { width: 100%; border-radius: 20px; height: 3em; background-color: #4A90E2; color: white; }
    .stDownloadButton>button { width: 100%; border-radius: 20px; background-color: #2ECC71; color: white; }
    [data-testid="stExpander"] { border: 1px solid #e6e6e6; border-radius: 10px; background-color: #fcfcfc; }
    </style>
    """, unsafe_allow_html=True)

with st.sidebar:
    st.title("⚙️ 配置参数")
    st.divider()
    st.info("💡 **说明**：\n1. Excel 第一个 Sheet 需包含 `placeholder` 和 `value` 列。\n2. 第二个 Sheet 需包含 `数据库`、`文献信息` 等。")
    
    st.subheader("📊 表格顺序设置")
    st.caption("对应模板中表格出现的先后顺序 (0开始)")
    idx_wanfang = st.number_input("万方数据表索引", value=3)
    idx_cnki = st.number_input("知网数据表索引", value=4)
    idx_pubmed = st.number_input("Pubmed数据表索引", value=5)
    idx_embase = st.number_input("Embase数据表索引", value=6)

st.title("📄 PRER报告自动生成工具（测试）")
st.write("上传您的数据，3s完成格式化报告生成。")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1️⃣ 模板文件")
    template_file = st.file_uploader("上传 Word 模板 (.docx)", type=["docx"])

with col2:
    st.subheader("2️⃣ 数据源")
    excel_file = st.file_uploader("上传 Excel 数据 (.xlsx)", type=["xlsx"])

if excel_file:
    with st.expander("🔍 预览 Excel 数据"):
        try:
            xl = pd.ExcelFile(excel_file, engine='openpyxl')
            sheet = st.selectbox("选择要查看的分页", xl.sheet_names)
            st.dataframe(pd.read_excel(excel_file, sheet_name=sheet).head(5))
        except Exception as e:
            st.error(f"Excel 读取失败: {e}")

st.divider()

if st.button("🚀 开始生成报告"):
    if template_file and excel_file:
        try:
            with st.status("正在流水线处理...", expanded=True) as status:
                # 1. 数据读取
                st.write("读取并解析 Excel...")
                xl = pd.ExcelFile(excel_file, engine='openpyxl')
                p_df = pd.read_excel(excel_file, sheet_name=xl.sheet_names[0])
                l_df = pd.read_excel(excel_file, sheet_name=xl.sheet_names[1])
                
                # 构造替换字典
                replacements = {}
                for _, row in p_df.iterrows():
                    key = str(row.get("placeholder", "")).strip()
                    if not key: continue
                    if not key.startswith("{"): key = "{" + key
                    if not key.endswith("}"): key = key + "}"
                    replacements[key] = format_date(row.get("value", ""))

                # 文献分类逻辑
                st.write("分类文献数据库...")
                databases = {}
                for _, row in l_df.iterrows():
                    db_raw = str(row.get("数据库", "")).strip().upper()
                    if "万方" in db_raw: db = "万方"
                    elif any(k in db_raw for k in ["知网", "CNKI"]): db = "中国知网"
                    elif "PUBMED" in db_raw: db = "Pubmed"
                    elif "EMBASE" in db_raw: db = "Embase"
                    else: continue
                    
                    record = [str(row.get("文献编号", "")), str(row.get("文献信息", "")), 
                              str(row.get("检索说明", "")), str(row.get("排除理由", "")), 
                              str(row.get("备注", ""))]
                    databases.setdefault(db, []).append(record)

                total = sum(len(v) for v in databases.values())
                replacements["{文献量}"] = str(total)
                replacements["{总纳入文献量}"] = str(total)

                # 2. 文档替换
                st.write("执行正文与页眉替换...")
                doc = Document(template_file)
                replace_all_content_keep_style(doc, replacements)
                replace_header_keep_format(doc, replacements)

                # 3. 表格填充
                st.write("填充明细表格数据...")
                db_map = {idx_wanfang: "万方", idx_cnki: "中国知网", idx_pubmed: "Pubmed", idx_embase: "Embase"}
                for idx, db_name in db_map.items():
                    if idx < len(doc.tables):
                        fill_specific_table(doc.tables[idx], databases.get(db_name, []))

                # 4. 导出
                out_io = io.BytesIO()
                doc.save(out_io)
                out_io.seek(0)
                status.update(label="✅ 报告生成完毕！", state="complete")

            st.balloons()
            st.download_button(
                label="📥 点击下载生成的报告",
                data=out_io,
                file_name=f"自动生成报告_{datetime.now().strftime('%m%d%H%M')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"❌ 处理过程中出错: {e}")
            st.info("请检查 Excel 列名是否包含：placeholder, value, 数据库, 文献编号, 文献信息 等。")
    else:
        st.warning("⚠️ 请先上传模板和数据文件。")

# --- 版权信息 ---
st.divider()
st.caption("© CER中心 自动化排版工作室 | Seven")
