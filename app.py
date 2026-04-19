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
# 样式优化（清爽 + SaaS风格）
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

.card {
    padding: 16px;
    border-radius: 12px;
    background-color: #f7f9fc;
    border: 1px solid #e6eaf0;
    margin-bottom: 10px;
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
# 标题区
# =========================
st.markdown("<div class='main-title'>📄 PRER 自动报告生成系统</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-title'>基于多源文献数据自动生成结构化医疗报告</div>", unsafe_allow_html=True)

st.divider()

# =========================
# 侧边栏：流程说明
# =========================
with st.sidebar:
    st.header("⚙️ 使用流程")

    st.markdown("""
    1️⃣ 上传产品信息表  
    2️⃣ 上传文献筛选表  
    3️⃣ 上传Word模板  
    4️⃣ 可选上传分析数据  
    5️⃣ 点击生成报告  
    """)

    st.divider()

    st.info("💡 必填：产品信息 + 文献筛选 + 模板")

# =========================
# 上传区（结构化分组）
# =========================
st.subheader("📂 数据上传区")

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("#### 📌 核心数据（必填）")
    product_info = st.file_uploader("产品信息表", type=["xlsx"])
    screening = st.file_uploader("文献筛选表", type=["xlsx"])
    template = st.file_uploader("Word模板", type=["docx"])

with col2:
    st.markdown("#### 📊 分析数据（可选）")
    analysis = st.file_uploader("文献数据分析表", type=["xlsx"])
    efficacy = st.file_uploader("有效性结果表", type=["xlsx"])

with col3:
    st.markdown("#### 🧪 安全数据（可选）")
    safety = st.file_uploader("安全性结果表", type=["xlsx"])

# =========================
# 数据状态可视化卡片
# =========================
st.subheader("📊 数据状态")

c1, c2, c3 = st.columns(3)

c1.metric("产品信息表", "✔ 已上传" if product_info else "❌ 未上传")
c2.metric("文献筛选表", "✔ 已上传" if screening else "❌ 未上传")
c3.metric("Word模板", "✔ 已上传" if template else "❌ 未上传")

# =========================
# 执行按钮区（重点突出）
# =========================
st.markdown("---")

st.markdown("<div class='big-button'>", unsafe_allow_html=True)

if st.button("🚀 开始生成PRER报告"):
    if not product_info or not screening or not template:
        st.error("❌ 请先上传：产品信息表 + 文献筛选表 + Word模板")
    else:
        with st.spinner("🧠 AI正在生成结构化报告，请稍候..."):

            file_map = {
                "product_info": product_info,
                "screening": screening,
                "analysis": analysis,
                "efficacy": efficacy,
                "safety": safety,
            }

            output = generate_report(file_map, template)

        st.success("✅ 报告生成成功！")

        st.markdown(
            """
            <div class="success-box">
            🎉 文件已准备完成，可立即下载
            </div>
            """,
            unsafe_allow_html=True
        )

        st.download_button(
            label="📥 下载PRER报告",
            data=output,
            file_name=f"PRER报告_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

st.markdown("</div>", unsafe_allow_html=True)

# =========================
# 底部信息
# =========================
st.divider()
st.caption("© CER中心 | PRER自动化报告系统 | Seven")```

---

# 🎯 你这个版本相比原版的提升

## ✔ 1. 信息结构升级
原来：一堆 uploader  
现在：  
👉 核心数据 / 分析数据 / 安全数据 分区

---

## ✔ 2. 可视化状态（关键提升）
```python
st.metric()
