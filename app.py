import streamlit as st
from core import generate_report

st.set_page_config(page_title="PRER报告生成系统", layout="wide")

st.title("📄 PRER 自动报告生成系统")

st.markdown("上传五个Excel和Word模板，自动生成报告")

# =========================
# 文件上传
# =========================
product_info = st.file_uploader("产品信息表", type=["xlsx"])
screening = st.file_uploader("文献筛选表", type=["xlsx"])
analysis = st.file_uploader("文献数据分析表", type=["xlsx"])
efficacy = st.file_uploader("有效性结果表", type=["xlsx"])
safety = st.file_uploader("安全性结果表", type=["xlsx"])
template = st.file_uploader("Word模板", type=["docx"])

# =========================
# 按钮
# =========================
if st.button("🚀 生成报告"):

    if not product_info or not screening or not template:
        st.error("❌ 必须上传：产品信息表 + 文献筛选表 + 模板")
    else:
        with st.spinner("正在生成报告..."):

            file_map = {
                "product_info": product_info,
                "screening": screening,
                "analysis": analysis,
                "efficacy": efficacy,
                "safety": safety,
            }

            output = generate_report(file_map, template)

            st.success("✅ 生成完成！")

            st.download_button(
                label="📥 下载报告",
                data=output,
                file_name="PRER报告.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

# --- 版权信息 ---
st.divider()
st.caption("© CER中心 自动化排版工作室 | Seven")
