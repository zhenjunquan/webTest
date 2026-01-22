import streamlit as st
import subprocess
import os
import shutil
import tarfile
import urllib.request
import tempfile
import sys

# ==========================================
# 🛠️ 1. 环境自动配置 (保持不变，确保云端可用)
# ==========================================
def install_linux_tools():
    base_dir = os.getcwd()
    bin_dir = os.path.join(base_dir, "bin")
    pandoc_exe = os.path.join(bin_dir, "pandoc")
    crossref_exe = os.path.join(bin_dir, "pandoc-crossref")

    if os.path.exists(pandoc_exe) and os.path.exists(crossref_exe):
        return bin_dir

    st.toast("正在初始化 Pandoc 环境 (首次运行需 30s)...", icon="🚀")
    if not os.path.exists(bin_dir): os.makedirs(bin_dir)

    # 下载 Pandoc (Linux)
    PANDOC_VER = "3.1.12.3"
    p_url = f"https://github.com/jgm/pandoc/releases/download/{PANDOC_VER}/pandoc-{PANDOC_VER}-linux-amd64.tar.gz"
    try:
        t_path, _ = urllib.request.urlretrieve(p_url)
        with tarfile.open(t_path) as t:
            for m in t.getmembers():
                if m.name.endswith("bin/pandoc"):
                    m.name = "pandoc"; t.extract(m, bin_dir)
    except Exception as e: st.error(f"Pandoc 下载失败: {e}")

    # 下载 Crossref (Linux)
    CROSSREF_VER = "0.3.17.1a"
    c_url = f"https://github.com/lierdakil/pandoc-crossref/releases/download/v{CROSSREF_VER}/pandoc-crossref-Linux.tar.xz"
    try:
        t_path, _ = urllib.request.urlretrieve(c_url)
        with tarfile.open(t_path, "r:xz") as t:
            for m in t.getmembers():
                if m.name.endswith("pandoc-crossref"):
                    m.name = "pandoc-crossref"; t.extract(m, bin_dir)
    except Exception as e: st.error(f"Crossref 下载失败: {e}")

    subprocess.run(["chmod", "+x", pandoc_exe])
    subprocess.run(["chmod", "+x", crossref_exe])
    return bin_dir

# 环境检测逻辑
if sys.platform.startswith("linux"):
    local_bin = install_linux_tools()
    os.environ["PATH"] = local_bin + os.pathsep + os.environ["PATH"]
    CROSSREF_CMD = os.path.join(local_bin, "pandoc-crossref")
else:
    # 本地 Windows 开发环境，假设已安装 scoop
    CROSSREF_CMD = "pandoc-crossref"

# ==========================================
# 🎨 2. 界面与逻辑
# ==========================================

st.set_page_config(page_title="Pandoc Pro Converter", layout="wide", page_icon="📝")
st.title("📝 Markdown 转 Word (全功能版)")

# 默认 YAML (保持你的配置)
DEFAULT_YAML = """---
lang: en
chapters: true
linkReferences: true
chapDelim: "-"
figPrefix: 
figureTemplate: 图 $$i$$. $$t$$
tblPrefix: 
tableTemplate: Table $$i$$ $$t$$
secPrefix: 节
reference-section-title: 参考文献
reference-section-number: false
link-citations: true
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

# --- 侧边栏：配置区 ---
with st.sidebar:
    st.header("⚙️ 基础设置")
    template_file = st.file_uploader("1. 样式模板 (templates.docx)", type=["docx"])
    
    st.header("🛠️ 常用命令开关")
    opt_toc = st.checkbox("生成目录 (--toc)", value=False)
    opt_number = st.checkbox("章节编号 (--number-sections)", value=True)
    opt_citeproc = st.checkbox("处理参考文献 (--citeproc)", value=False, help="如果你使用了 bib 文件，请勾选此项")
    
    st.header("📝 元数据配置")
    yaml_content = st.text_area("编辑 Meta.yaml", value=DEFAULT_YAML, height=300)

# --- 主界面：文件与输出 ---
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. 上传文件")
    source_file = st.file_uploader("上传论文 Markdown (.md)", type=["md"])
    
    st.subheader("2. 上传图片 (可选)")
    uploaded_images = st.file_uploader("选择文中引用的所有图片", accept_multiple_files=True, type=["png", "jpg", "jpeg", "svg", "pdf"])
    img_folder_name = st.text_input("Markdown 中的图片文件夹名称", value="assets", 
                                    help="例如你的 MD 里写的是 ![](assets/pic.png)，这里就填 assets。如果写的是 ![](pic.png)，这里留空。")

with col2:
    st.subheader("3. 输出设置")
    output_filename = st.text_input("下载文件名", value="paper_final", help="不需要加 .docx 后缀")
    if not output_filename.endswith(".docx"):
        output_filename += ".docx"
    
    st.write("---")
    start_btn = st.button("🚀 开始转换", type="primary", use_container_width=True)

# --- 转换核心逻辑 ---
if start_btn and source_file:
    # 环境检查
    try:
        subprocess.run(["pandoc", "-v"], capture_output=True)
    except FileNotFoundError:
        st.error("❌ 环境未就绪，请等待 Pandoc 下载完成或刷新页面。")
        st.stop()

    with st.spinner("正在构建文档..."):
        with tempfile.TemporaryDirectory() as temp_dir:
            # 1. 保存 MD 源文件
            input_path = os.path.join(temp_dir, "paper.md")
            with open(input_path, "wb") as f:
                f.write(source_file.getvalue())
            
            # 2. 保存 YAML
            yaml_path = os.path.join(temp_dir, "meta.yaml")
            with open(yaml_path, "w", encoding="utf-8") as f:
                f.write(yaml_content)
            
            # 3. 处理图片 (关键步骤)
            if uploaded_images:
                # 确定图片存放路径
                if img_folder_name.strip():
                    img_save_dir = os.path.join(temp_dir, img_folder_name)
                    if not os.path.exists(img_save_dir):
                        os.makedirs(img_save_dir)
                else:
                    img_save_dir = temp_dir # 直接放在根目录
                
                # 保存每一张图片
                for img_file in uploaded_images:
                    img_path = os.path.join(img_save_dir, img_file.name)
                    with open(img_path, "wb") as f:
                        f.write(img_file.getvalue())
                
                st.toast(f"已处理 {len(uploaded_images)} 张图片", icon="🖼️")

            # 4. 保存模板
            cmd_template = []
            if template_file:
                tpl_path = os.path.join(temp_dir, "templates.docx")
                with open(tpl_path, "wb") as f:
                    f.write(template_file.getvalue())
                cmd_template = [f"--reference-doc={tpl_path}"]

            # 5. 组装命令
            # 基础命令
            cmd = [
                "pandoc", 
                input_path, 
                f"--metadata-file={yaml_path}", 
                "--filter", CROSSREF_CMD, 
                "-o", "output.docx"
            ]
            
            # 加入用户选定的常用命令
            if opt_toc: cmd.append("--toc")
            if opt_number: cmd.append("--number-sections")
            if opt_citeproc: cmd.append("--citeproc")
            
            # 加入模板参数
            cmd.extend(cmd_template)

            # 6. 执行
            process = subprocess.run(cmd, cwd=temp_dir, capture_output=True, text=True)

            if process.returncode == 0:
                out_path = os.path.join(temp_dir, "output.docx")
                with open(out_path, "rb") as f:
                    st.success("✅ 转换成功！")
                    st.download_button(
                        label=f"📥 下载 {output_filename}",
                        data=f.read(),
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary"
                    )
            else:
                st.error("❌ 转换失败")
                with st.expander("查看详细错误日志"):
                    st.code(process.stderr)
