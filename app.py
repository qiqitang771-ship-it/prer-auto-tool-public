import os
import streamlit as st
from core import generate_report
from datetime import datetime

# =========================
# 页面配置
# =========================
st.set_page_config(
    page_title="PRER报告生成系统",
    layout="wide",
    page_icon="📄"
)

# =========================
# 样式优化
# =========================
st.markdown("""
<style>
.main-title {
    font-size: 28px;
    font-weight: 700;
    margin-bottom: 5px;
}

.sub-title {
    color: #666;
    margin-bottom: 20px;
}

.big-button button {
    width: 100%;
    height: 3.2em;
    font-size: 16px;
    font-weight: 600;
    border-radius: 12px;
    background: linear-gradient(90deg, #4A90E2, #357ABD);
    color: white;
}

.success-box {
    padding: 12px;
    background-color: #e8f8f0;
    border-radius: 10px;
    border: 1px solid #b7e4c7;
}
</style>
""", unsafe_allow_html=True)

# =========================
# 标题
# =========================
st.markdown("<div class='main-title'>📄 PRER 自动报告生成系统</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-title'>基于文献数据自动生成结构化报告</div>", unsafe_allow_html=True)

st.divider()

# =========================
# sidebar
# =========================
with st.sidebar:
    st.header("⚙️ 使用流程")
    st.markdown("""
    1️⃣ 上传产品信息表  
    2️⃣ 上传文献筛选表  
    3️⃣ 选择固定Word模板  
    4️⃣ 可选上传分析数据  
    5️⃣ 点击生成报告  
    """)
    st.divider()
    st.info("💡 必填：产品信息 + 文献筛选 + 模板")

# =========================
# 上传区（修复版）
# =========================
st.subheader("📂 数据上传区")

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("#### 📌 核心数据（必填）")
    product_info = st.file_uploader("产品信息表", type=["xlsx"], key="product_info")
    screening = st.file_uploader("文献筛选表", type=["xlsx"], key="screening")

template_options = {
    "模板1 - 标准模板": "template-2个数据库.docx",
    "模板2 - 简洁模板": "template-4个数据库.docx",
}

with col2:
    st.markdown("#### 📄 报告模板（必填）")
    template_choice = st.selectbox(
        "请选择模板",
        options=list(template_options.keys()),
        help="请选择固定模板1或模板2",
        key="template_select"
    )
    selected_template_path = template_options.get(template_choice)
    template_exists = os.path.exists(selected_template_path)

    if template_exists:
        st.success(f"已选择：{template_choice}")
    else:
        st.error(f"模板文件未找到：{selected_template_path}")

with col3:
    st.markdown("#### 📊 文献分析数据（可选）")
    analysis = st.file_uploader("文献数据分析表", type=["xlsx"], key="analysis")

# =========================
# 数据状态
# =========================
st.subheader("📊 数据状态")

c1, c2, c3 = st.columns(3)

c1.metric("产品信息表", "✔ 已上传" if product_info else "❌ 未上传")
c2.metric("文献筛选表", "✔ 已上传" if screening else "❌ 未上传")
c3.metric("Word模板", "✔ 已选择" if template_exists else "❌ 未找到")

# =========================
# 执行区
# =========================
st.markdown("---")

st.markdown("<div class='big-button'>", unsafe_allow_html=True)

if st.button("🚀 开始生成PRER报告"):

    if not product_info or not screening or not template_exists:
        st.error("❌ 请先准备：产品信息表 + 文献筛选表，并确保选择模板")

    else:
        with st.spinner("🧠 正在生成结构化报告，请稍候..."):
            file_map = {
                "product_info": product_info,
                "screening": screening,
                "analysis": analysis,
            }

            with open(selected_template_path, "rb") as template_file:
                output = generate_report(file_map, template_file)

        st.success("✅ 报告生成成功！")

        st.markdown("""
        <div class="success-box">
        🎉 文件已准备完成，可立即下载
        </div>
        """, unsafe_allow_html=True)

        filename = f"PRER报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"

        st.download_button(
            label="📥 下载PRER报告",
            data=output,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

st.markdown("</div>", unsafe_allow_html=True)

# =========================
# footer
# =========================
st.divider()
st.caption("© CER中心 | PRER自动化报告系统 | Seven")
