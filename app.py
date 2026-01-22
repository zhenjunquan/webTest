import streamlit as st
import subprocess
import os
import tempfile
import shutil

# ================= 配置区域 =================
st.set_page_config(page_title="Pandoc 论文转换器", layout="wide")

st.title("📄 学术 Markdown 转 Word (Pandoc Pro)")
st.markdown("支持 `pandoc-crossref` 交叉引用与 `templates.docx` 样式定制")

# 默认的 YAML 配置 (你提供的内容)
DEFAULT_YAML = """---
# ============ pandoc-crossref 基础配置 ============
lang: en
chapters: true
linkReferences: true
chapDelim: "-"

# 图标题设置
figPrefix:   #引用图
figureTemplate:  图 $$i$$.  $$t$$  #图标题格式

# 表标题设置
tblPrefix:   #引用表
tableTemplate: Table $$i$$  $$t$$  #表标题格式

# 引用表题
secPrefix: 节

#参考文献
reference-section-title: 参考文献
reference-section-number: false
link-citations: true

# 公式相关
eqnos: true
eqnPrefix: 式
autoEqnLabels: true
tableEqns: true
eqnBlockTemplate: |
   `<w:pPr><w:jc w:val="center"/><w:spacing w:line="400" w:lineRule="atLeast"/><w:tabs><w:tab w:val="center" w:leader="none" w:pos="4478" /><w:tab w:val="right" w:leader="none" w:pos="10433" /></w:tabs></w:pPr><w:r><w:tab /></w:r>`{=openxml} $$t$$ `<w:r><w:tab /></w:r>`{=openxml} $$i$$

eqnBlockInlineMath: true
equationNumberTeX: \\\\tag
eqnIndexTemplate: ($$i$$)
eqnPrefixTemplate: 式($$i$$)
---"""

# ================= 侧边栏：配置与模板 =================
with st.sidebar:
    st.header("⚙️ 编译设置")
    
    # 1. 上传样式模板 (可选)
    template_file = st.file_uploader("上传样式模板 (templates.docx)", type=["docx"])
    if template_file:
        st.success(f"已加载模板: {template_file.name}")
    else:
        st.info("未上传模板，将使用 Pandoc 默认样式")

    # 2. 编辑 Metadata (Yaml)
    st.subheader("元数据配置 (meta.yaml)")
    yaml_content = st.text_area("可在此处直接修改配置", value=DEFAULT_YAML, height=400)

# ================= 主区域：转换逻辑 =================
source_file = st.file_uploader("请上传论文 Markdown 文件 (.md)", type=["md"])

if source_file and st.button("开始转换 (Convert)", type="primary"):
    with st.spinner("正在调用 Pandoc 编译中..."):
        try:
            # 创建临时目录来存放所有文件
            with tempfile.TemporaryDirectory() as temp_dir:
                
                # 1. 保存 source.md
                input_path = os.path.join(temp_dir, "paper.md")
                with open(input_path, "wb") as f:
                    f.write(source_file.getvalue())
                
                # 2. 保存 meta.yaml
                yaml_path = os.path.join(temp_dir, "meta.yaml")
                with open(yaml_path, "w", encoding="utf-8") as f:
                    f.write(yaml_content)
                
                # 3. 构建 Pandoc 命令
                # 基础命令
                cmd = [
                    "pandoc", 
                    input_path, 
                    f"--metadata-file={yaml_path}", 
                    "--filter", "pandoc-crossref", 
                    "-o", "paper.docx"
                ]

                # 4. 处理模板 (如果上传了的话)
                if template_file:
                    template_path = os.path.join(temp_dir, "templates.docx")
                    with open(template_path, "wb") as f:
                        f.write(template_file.getvalue())
                    # 添加参数
                    cmd.extend([f"--reference-doc={template_path}"])
                
                # 5. 执行命令
                # 注意：cwd=temp_dir 保证了 pandoc 在临时目录运行，输出也在那里
                process = subprocess.run(cmd, cwd=temp_dir, capture_output=True, text=True)

                if process.returncode == 0:
                    output_path = os.path.join(temp_dir, "paper.docx")
                    with open(output_path, "rb") as f:
                        docx_data = f.read()
                    
                    st.success("✅ 转换成功！")
                    st.download_button(
                        label="📥 下载最终 Word 文档",
                        data=docx_data,
                        file_name="paper_final.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    st.error("❌ 转换失败")
                    st.error("错误详情:")
                    st.code(process.stderr)
                    
        except Exception as e:
            st.error(f"发生未知错误: {e}")

# 检查环境提示
try:
    subprocess.run(["pandoc", "-v"], stdout=subprocess.DEVNULL)
except FileNotFoundError:
    st.warning("⚠️ 警告：未检测到 Pandoc，请确保服务器已安装 pandoc。")

try:
    subprocess.run(["pandoc-crossref", "--version"], stdout=subprocess.DEVNULL)
except FileNotFoundError:
    st.warning("⚠️ 警告：未检测到 pandoc-crossref，交叉引用功能将失效。")
