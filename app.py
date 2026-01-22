import streamlit as st
import subprocess
import os
import shutil
import tarfile
import urllib.request
import tempfile
import sys
import zipfile

# ==========================================
# 🛠️ 1. 环境自动配置 (保持不变)
# ==========================================
def install_linux_tools():
    """云端自动安装 Pandoc 环境"""
    base_dir = os.getcwd()
    bin_dir = os.path.join(base_dir, "bin")
    pandoc_exe = os.path.join(bin_dir, "pandoc")
    crossref_exe = os.path.join(bin_dir, "pandoc-crossref")

    if os.path.exists(pandoc_exe) and os.path.exists(crossref_exe):
        return bin_dir

    st.toast("正在初始化环境...", icon="🚀")
    if not os.path.exists(bin_dir): os.makedirs(bin_dir)

    PANDOC_VER = "3.1.12.3"
    p_url = f"https://github.com/jgm/pandoc/releases/download/{PANDOC_VER}/pandoc-{PANDOC_VER}-linux-amd64.tar.gz"
    try:
        t_path, _ = urllib.request.urlretrieve(p_url)
        with tarfile.open(t_path) as t:
            for m in t.getmembers():
                if m.name.endswith("bin/pandoc"):
                    m.name = "pandoc"; t.extract(m, bin_dir)
    except: pass

    CROSSREF_VER = "0.3.17.1a"
    c_url = f"https://github.com/lierdakil/pandoc-crossref/releases/download/v{CROSSREF_VER}/pandoc-crossref-Linux.tar.xz"
    try:
        t_path, _ = urllib.request.urlretrieve(c_url)
        with tarfile.open(t_path, "r:xz") as t:
            for m in t.getmembers():
                if m.name.endswith("pandoc-crossref"):
                    m.name = "pandoc-crossref"; t.extract(m, bin_dir)
    except: pass

    subprocess.run(["chmod", "+x", pandoc_exe])
    subprocess.run(["chmod", "+x", crossref_exe])
    return bin_dir

# 环境检测
if sys.platform.startswith("linux"):
    local_bin = install_linux_tools()
    os.environ["PATH"] = local_bin + os.pathsep + os.environ["PATH"]
    CROSSREF_CMD = os.path.join(local_bin, "pandoc-crossref")
else:
    CROSSREF_CMD = "pandoc-crossref"

# ==========================================
# 🎨 2. 界面布局与逻辑
# ==========================================

st.set_page_config(page_title="Pandoc Pro", layout="wide", page_icon="📑")
st.title("📑 Markdown 转 Word (Zip版)")

# 默认配置
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

# --- 侧边栏 ---
with st.sidebar:
    st.header("⚙️ 1. 上传")
    st.info("💡 推荐：将 .md 和图片打包成 Zip 上传，可自动保持路径结构。")
    
    upload_type = st.radio("选择上传方式", ["上传 Zip 压缩包 (推荐)", "仅上传单个 MD 文件"])
    
    source_file = None
    if upload_type == "上传 Zip 压缩包 (推荐)":
        source_file = st.file_uploader("上传包含 MD 和图片的 Zip", type=["zip"])
    else:
        source_file = st.file_uploader("上传 Markdown 文件", type=["md"])
    
    st.header("🎨 2. 样式")
    template_file = st.file_uploader("样式模板 (templates.docx)", type=["docx"])
    
    st.header("🔧 3. 选项")
    opt_toc = st.checkbox("生成目录 (--toc)", False)
    opt_num = st.checkbox("章节编号 (--number-sections)", True)
    output_name = st.text_input("输出文件名", "paper_final")

# --- 主界面 ---
# 使用 Tabs 分割预览和配置
tab1, tab2 = st.tabs(["👁️ 内容预览 & 转换", "⚙️ Meta 配置"])

with tab2:
    yaml_content = st.text_area("编辑 YAML 配置", DEFAULT_YAML, height=400)

with tab1:
    if source_file:
        # 创建临时文件夹来解压或保存文件
        with tempfile.TemporaryDirectory() as temp_dir:
            md_path = ""
            
            # --- 核心逻辑：文件处理 ---
            if source_file.name.endswith('.zip'):
                # 1. 解压 Zip
                zip_path = os.path.join(temp_dir, "upload.zip")
                with open(zip_path, "wb") as f:
                    f.write(source_file.getvalue())
                
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                
                # 2. 自动寻找 .md 文件
                found_md = False
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        if file.endswith(".md"):
                            md_path = os.path.join(root, file)
                            found_md = True
                            break # 默认取第一个 md
                    if found_md: break
                
                if not found_md:
                    st.error("❌ Zip 包里没找到 .md 文件！")
                    st.stop()
            else:
                # 普通 MD 上传
                md_path = os.path.join(temp_dir, source_file.name)
                with open(md_path, "wb") as f:
                    f.write(source_file.getvalue())

            # --- 👁️ 功能：Markdown 预览 ---
            try:
                with open(md_path, "r", encoding="utf-8") as f:
                    md_content = f.read()
                
                st.subheader(f"📄 预览: {os.path.basename(md_path)}")
                with st.expander("点击展开/折叠 Markdown 内容预览", expanded=True):
                    st.markdown(md_content)
                    # st.text_area("源码预览", md_content, height=200) # 也可以用纯文本显示
            except Exception as e:
                st.warning(f"无法预览文件内容: {e}")

            # --- 转换按钮 ---
            st.write("---")
            if st.button("🚀 开始转换 Word", type="primary"):
                with st.spinner("正在调用 Pandoc 引擎..."):
                    # 写入 YAML
                    yaml_path = os.path.join(temp_dir, "meta.yaml")
                    with open(yaml_path, "w", encoding="utf-8") as f:
                        f.write(yaml_content)
                    
                    # 写入模板
                    cmd_template = []
                    if template_file:
                        tpl_path = os.path.join(temp_dir, "template.docx")
                        with open(tpl_path, "wb") as f:
                            f.write(template_file.getvalue())
                        cmd_template = [f"--reference-doc={tpl_path}"]

                    # 构建命令
                    # 注意：cwd=os.path.dirname(md_path) 确保 pandoc 在 md 文件所在的目录运行
                    # 这样 md 里的相对路径引用 (如 images/1.png) 才能生效
                    work_dir = os.path.dirname(md_path)
                    
                    cmd = [
                        "pandoc", 
                        md_path, 
                        f"--metadata-file={yaml_path}", 
                        "--filter", CROSSREF_CMD,
                        "-o", "output.docx"
                    ]
                    if opt_toc: cmd.append("--toc")
                    if opt_num: cmd.append("--number-sections")
                    cmd.extend(cmd_template)

                    # 执行
                    res = subprocess.run(cmd, cwd=work_dir, capture_output=True, text=True)

                    if res.returncode == 0:
                        out_path = os.path.join(work_dir, "output.docx")
                        with open(out_path, "rb") as f:
                            file_data = f.read()
                        
                        st.success("✅ 转换成功！")
                        
                        # --- 👁️ 功能：Word 简易信息预览 ---
                        # 浏览器无法直接预览 Word 内容，但我们可以显示文件信息
                        file_size = len(file_data) / 1024
                        st.info(f"生成文件大小: {file_size:.2f} KB")
                        
                        full_name = output_name if output_name.endswith(".docx") else output_name + ".docx"
                        st.download_button("📥 点击下载 Word 文档", file_data, full_name, type="primary")
                    else:
                        st.error("❌ 转换失败")
                        st.code(res.stderr)

    else:
        st.info("👈 请在左侧上传文件 (支持 .md 或包含资源的 .zip)")
