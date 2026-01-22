import streamlit as st
import pypandoc
import os
import tempfile

st.title("Markdown 转 Word 工具 📝")
st.write("上传你的 .md 文件，我会把它转换成 .docx 供你下载。")

# 1. 文件上传组件
uploaded_file = st.file_uploader("选择一个 Markdown 文件", type=["md"])

if uploaded_file is not None:
    # 创建临时文件来处理
    with tempfile.NamedTemporaryFile(delete=False, suffix=".md") as tmp_input:
        tmp_input.write(uploaded_file.getvalue())
        input_path = tmp_input.name

    output_path = input_path.replace(".md", ".docx")

    try:
        # 2. 调用 Pandoc 进行转换 (核心逻辑)
        # 这里的 outputfile 指定输出路径
        pypandoc.convert_file(input_path, 'docx', outputfile=output_path)
        
        # 3. 读取转换后的文件准备下载
        with open(output_path, "rb") as f:
            file_data = f.read()

        st.success("转换成功！")
        
        # 4. 下载按钮
        st.download_button(
            label="下载 Word 文档 (.docx)",
            data=file_data,
            file_name="converted_document.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    except Exception as e:
        st.error(f"转换出错: {e}")
        st.info("提示：如果遇到 Pandoc 错误，通常是因为环境里没装 Pandoc。但在 Streamlit Cloud 上我们会通过配置自动安装。")
        
    finally:
        # 清理临时文件
        if os.path.exists(input_path):
            os.remove(input_path)
        if os.path.exists(output_path):
            os.remove(output_path)
