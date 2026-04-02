import streamlit as st
import pandas as pd
from docx import Document
import io
import copy
from datetime import datetime
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import Pt
from docx.oxml.ns import qn

# --- 核心逻辑函数 (保持原代码功能并微调) ---

def format_date(value):
    if pd.isna(value): return ""
    if isinstance(value, (pd.Timestamp, datetime)):
        return f"{value.year}年{value.month}月{value.day}日"
    return str(value)

def set_dual_font_10pt(run, font_name='Arial', east_asia_font='宋体'):
    run.font.name = font_name
    run.font.size = Pt(10)
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), east_asia_font)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)

def replace_text_strict_format(paragraph, replacements):
    if not paragraph.runs: return
    full_text = "".join(run.text for run in paragraph.runs)
    new_text = full_text
    for k, v in replacements.items():
        new_text = new_text.replace(k, str(v))
    if new_text != full_text:
        paragraph.runs[0].text = new_text
        for i in range(1, len(paragraph.runs)):
            paragraph.runs[i].text = ""

def fill_specific_table(table, records):
    if not records or len(table.rows) < 2: return
    template_row = table.rows[1]
    while len(table.rows) > 2:
        table._tbl.remove(table.rows[-1]._tr)
    for idx, record in enumerate(records, start=1):
        new_row_xml = copy.deepcopy(template_row._tr)
        table._tbl.append(new_row_xml)
        row = table.rows[-1]
        # 文献编号, 文献信息, 检索说明, 排除理由, 备注
        row_data = [str(idx), record[1], record[2], record[3], record[4]]
        for i, val in enumerate(row_data):
            if i < len(row.cells):
                cell = row.cells[i]
                if not cell.paragraphs: cell.add_paragraph()
                p = cell.paragraphs[0]
                p.text = ""
                run = p.add_run(str(val))
                set_dual_font_10pt(run)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    table._tbl.remove(template_row._tr)

# --- Streamlit 页面配置 ---

st.set_page_config(page_title="PRER 报告自动化工具", layout="wide", page_icon="📝")

# 自定义 CSS 样式
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    div.stButton > button:first-child { background-color: #007bff; color: white; border-radius: 5px; width: 100%; height: 3em; }
    </style>
    """, unsafe_allow_html=True)

st.title("📝 PRER 报告自动化生成工具")
st.caption("基于 Python-Docx 和 Streamlit 构建的智能排版引擎")

# --- 侧边栏：灵活配置 ---
with st.sidebar:
    st.header("⚙️ 引擎配置")
    with st.expander("表格索引设置", expanded=False):
        idx_wf = st.number_input("万方表格索引", value=3)
        idx_cnki = st.number_input("知网表格索引", value=4)
        idx_pub = st.number_input("Pubmed表格索引", value=5)
        idx_em = st.number_input("Embase表格索引", value=6)
    
    with st.expander("数据识别关键词", expanded=False):
        kw_wf = st.text_input("万方关键词", "万方")
        kw_cnki = st.text_input("知网关键词", "知网,CNKI")
        kw_pub = st.text_input("Pubmed关键词", "PUBMED")
        kw_em = st.text_input("Embase关键词", "EMBASE")

    st.divider()
    st.help("提示：索引指的是 Word 文档中第几个表格（从0开始计数）。如果报告生成的表格位置不对，请调整上方索引。")

# --- 主界面：任务分栏 ---
tab1, tab2 = st.tabs(["📤 上传与预览", "🛠️ 执行生成"])

with tab1:
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("1. 数据源 (Excel)")
        excel_file = st.file_uploader("选择数据文件", type=["xlsx"], help="需包含占位符Sheet和文献明细Sheet")
    with c2:
        st.subheader("2. 报告模板 (Word)")
        word_file = st.file_uploader("选择模板文件", type=["docx"], help="请确保包含 {标签} 占位符")

    if excel_file:
        st.divider()
        st.subheader("📊 数据概览")
        xl = pd.ExcelFile(excel_file)
        sheet_choice = st.selectbox("查看表格数据", xl.sheet_names)
        preview_df = pd.read_excel(excel_file, sheet_name=sheet_choice)
        st.dataframe(preview_df.head(5), use_container_width=True)

with tab2:
    if not excel_file or not word_file:
        st.info("请先在 '上传与预览' 标签页中上传必要文件。")
    else:
        st.subheader("⚙️ 处理中心")
        if st.button("🔥 启动自动化流程"):
            try:
                # 1. 数据解析阶段
                excel_obj = pd.ExcelFile(excel_file)
                df_p = pd.read_excel(excel_file, sheet_name=excel_obj.sheet_names[0])
                df_l = pd.read_excel(excel_file, sheet_name=excel_obj.sheet_names[1])
                
                # 占位符转换
                replacements = {}
                for _, row in df_p.iterrows():
                    key = str(row["placeholder"]).strip()
                    if not key.startswith("{"): key = "{" + key
                    if not key.endswith("}"): key = key + "}"
                    replacements[key] = format_date(row["value"])

                # 文献分类逻辑
                databases = {"WF": [], "CNKI": [], "PUB": [], "EM": []}
                for _, row in df_l.iterrows():
                    db_raw = str(row.get("数据库", "")).strip().upper()
                    record = [str(row.get("文献编号", "")), str(row.get("文献信息", "")),
                              str(row.get("检索说明", "")), str(row.get("排除理由", "")),
                              str(row.get("备注", ""))]
                    
                    if any(k in db_raw for k in kw_wf.split(",")): databases["WF"].append(record)
                    elif any(k in db_raw for k in kw_cnki.upper().split(",")): databases["CNKI"].append(record)
                    elif any(k in db_raw for k in kw_pub.split(",")): databases["PUB"].append(record)
                    elif any(k in db_raw for k in kw_em.split(",")): databases["EM"].append(record)

                total_count = sum(len(v) for v in databases.values())
                replacements["{文献量}"] = str(total_count)
                replacements["{总纳入文献量}"] = str(total_count)

                # 2. 视觉反馈：统计指标
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("万方", len(databases["WF"]))
                m2.metric("知网", len(databases["CNKI"]))
                m3.metric("Pubmed", len(databases["PUB"]))
                m4.metric("Embase", len(databases["EM"]))

                # 3. Word 文档处理阶段
                with st.status("正在进行排版重构...", expanded=True) as status:
                    doc = Document(word_file)
                    
                    st.write("📝 替换正文和页眉标签...")
                    for p in doc.paragraphs: replace_text_strict_format(p, replacements)
                    for sec in doc.sections:
                        for hp in sec.header.paragraphs: replace_text_strict_format(hp, replacements)
                    
                    st.write("📊 正在注入文献明细表...")
                    tbs = doc.tables
                    mapping = [(idx_wf, "WF"), (idx_cnki, "CNKI"), (idx_pub, "PUB"), (idx_em, "EM")]
                    for idx, key in mapping:
                        if len(tbs) > idx:
                            fill_specific_table(tbs[idx], databases[key])
                    
                    st.write("🖼️ 处理文本框元数据...")
                    for node in doc.element.xpath("//w:t"):
                        if node.text:
                            for k, v in replacements.items(): node.text = node.text.replace(k, str(v))
                    
                    status.update(label="处理完成！", state="complete", expanded=False)

                # 4. 文件导出
                out_io = io.BytesIO()
                doc.save(out_io)
                out_io.seek(0)

                st.balloons()
                st.success(f"🎉 报告生成完毕！总计处理文献：{total_count} 篇")
                
                st.download_button(
                    label="💾 点击下载最终报告",
                    data=out_io,
                    file_name=f"PRER_Report_{datetime.now().strftime('%m%d_%H%M')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"处理失败，原因：{str(e)}")

# --- 版权信息 ---
st.divider()
st.caption("© CER中心 自动化排版工作室 | Seven")
