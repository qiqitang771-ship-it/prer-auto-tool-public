import streamlit as st
import os
import copy
import pandas as pd
from datetime import datetime
from docx import Document
import io

# --- 复用你之前的核心逻辑函数 ---
def format_date(value):
    if pd.isna(value): return ""
    if isinstance(value, (pd.Timestamp, datetime)):
        return f"{value.year}年{value.month}月{value.day}日"
    return str(value)

def replace_text_keep_format(paragraph, replacements):
    if not paragraph.runs: return
    full_text = "".join(run.text for run in paragraph.runs)
    new_text = full_text
    for k, v in replacements.items():
        if k in new_text:
            new_text = new_text.replace(k, str(v))
    if new_text != full_text:
        for i in range(1, len(paragraph.runs)):
            paragraph.runs[i].text = ""
        paragraph.runs[0].text = new_text

def replace_all_content(doc, replacements):
    for p in doc.paragraphs: replace_text_keep_format(p, replacements)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs: replace_text_keep_format(p, replacements)
    for node in doc.element.xpath("//w:t"):
        if node.text:
            txt = node.text
            for k, v in replacements.items():
                if k in txt: txt = txt.replace(k, str(v))
            node.text = txt

def fill_table_by_index(table, records, start_index=1):
    if not records or len(table.rows) < 2: return
    template_row = table.rows[1]
    while len(table.rows) > 2:
        table._tbl.remove(table.rows[-1]._tr)
    for idx, record in enumerate(records, start=start_index):
        new_row_xml = copy.deepcopy(template_row._tr)
        table._tbl.append(new_row_xml)
        row = table.rows[-1]
        try: row.cells[0].text = str(idx)
        except: pass
        values = [record[1], record[2], record[3], record[4]]
        for i, val in enumerate(values):
            col_index = i + 1
            if col_index < len(row.cells): row.cells[col_index].text = str(val)
    table._tbl.remove(template_row._tr)

# --- Streamlit 界面部分 ---
st.set_page_config(page_title="PRER 自动化排版系统", layout="centered")

st.title("📄 PRER 文献报告自动化工具")
st.markdown("上传 Excel 数据和 Word 模板，一键生成替换后的报告。")

with st.sidebar:
    st.header("配置参数")
    wf_index = st.number_input("万方表格索引 (默认3)", value=3)
    zw_index = st.number_input("知网表格索引 (默认4)", value=4)

col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("1. 上传 Excel 数据 (report_data.xlsx)", type=["xlsx"])
with col2:
    uploaded_word = st.file_uploader("2. 上传 Word 模板 (tmplate.docx)", type=["docx"])

if st.button("🚀 开始生成报告"):
    if uploaded_excel and uploaded_word:
        try:
            # 读取 Excel
            excel = pd.ExcelFile(uploaded_excel)
            p_sheet = excel.sheet_names[0]
            l_sheet = excel.sheet_names[1]
            
            # 加载占位符
            df_p = pd.read_excel(uploaded_excel, sheet_name=p_sheet)
            replacements = {}
            for _, row in df_p.iterrows():
                key = str(row["placeholder"]).strip()
                if not key.startswith("{"): key = "{" + key
                if not key.endswith("}"): key = key + "}"
                replacements[key] = format_date(row["value"])
            
            # 加载文献
            df_l = pd.read_excel(uploaded_excel, sheet_name=l_sheet)
            databases = {}
            for _, row in df_l.iterrows():
                db = str(row.get("数据库", "")).strip()
                if not db: continue
                record = [str(row.get("文献编号", "")), str(row.get("文献信息", "")), 
                          str(row.get("检索说明", "")), str(row.get("排除理由", "")), str(row.get("备注", ""))]
                databases.setdefault(db, []).append(record)
            
            # 计算总数
            wf_list = databases.get("万方", [])
            zw_list = databases.get("中国知网", []) or databases.get("CNKI", [])
            total = len(wf_list) + len(zw_list)
            replacements["{文献量}"] = str(total)
            replacements["{总纳入文献量}"] = str(total)

            # 处理 Word
            doc = Document(uploaded_word)
            replace_all_content(doc, replacements)
            
            tables = doc.tables
            fill_table_by_index(tables[wf_index], wf_list)
            fill_table_by_index(tables[zw_index], zw_list)
            
            # 保存到内存
            output = io.BytesIO()
            doc.save(output)
            output.seek(0)
            
            st.success("✅ 报告处理成功！")
            st.download_button(
                label="📥 点击下载最终报告",
                data=output,
                file_name="最终生成报告.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"处理过程中出错: {e}")
    else:
        st.warning("请先上传完整的 Excel 和 Word 文件。")