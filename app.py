import streamlit as st
import subprocess
import os
import tarfile
import urllib.request
import tempfile
import sys
import zipfile
import base64

# ==========================================
# 🛠️ 1. 环境配置 (Pandoc + LibreOffice检测)
# ==========================================
def install_linux_tools():
    """云端自动安装 Pandoc 环境"""
    base_dir = os.getcwd()
    bin_dir = os.path.join(base_dir, "bin")
    pandoc_exe = os.path.join(bin_dir, "pandoc")
    crossref_exe = os.path.join(bin_dir, "pandoc-crossref")

    if os.path.exists(pandoc_exe) and os.path.exists(crossref_exe):
        return bin_dir

    st.toast("正在初始化 Pandoc 环境...", icon="🚀")
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

# 环境初始化
if sys.platform.startswith("linux"):
    local_bin = install_linux_tools()
    os.environ["PATH"] = local_bin + os.pathsep + os.environ["PATH"]
    CROSSREF_CMD = os.path.join(local_bin, "pandoc-crossref")
else:
    CROSSREF_CMD = "pandoc-crossref"

# 检测 LibreOffice 是否可用
def check_libreoffice():
    """检测能否把 Word 转 PDF"""
    try:
        # 尝试调用 libreoffice (linux) 或 soffice (windows)
        cmd = "libreoffice" if sys.platform.startswith("linux") else "soffice"
        subprocess.run([cmd, "--version"], capture_output=True)
        return True, cmd
    except:
        return False, None

HAS_LO, LO_CMD = check_libreoffice()

# ==========================================
# 📂 2. 文件处理核心
# ==========================================
def unpack_and_find_md(upload_file, temp_dir):
    if upload_file.name.endswith('.zip'):
        zip_path = os.path.join(temp_dir, "upload.zip")
        with open(zip_path, "wb") as f:
            f.write(upload_file.getvalue())
        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
        except Exception as e:
            return None, None, f"解压失败: {e}"
    else:
        single_path = os.path.join(temp_dir, upload_file.name)
        with open(single_path, "wb") as f:
            f.write(upload_file.getvalue())

    for root, dirs, files in os.walk(temp_dir):
        for file in files:
            if file.endswith(".md") and not file.startswith("__"):
                return os.path.join(root, file), root, None
    
    return None, None, "未找到 .md 文件"

def convert_to_docx(md_path, work_dir, yaml_content, template_file, opt_toc, opt_num):
    """Pandoc 核心转换逻辑：MD -> DOCX"""
    # 写入配置
    yaml_path = os.path.join(work_dir, "meta.yaml")
    with open(yaml_path, "w", encoding="utf-8") as f: f.write(yaml_content)
    
    # 写入模板
    cmd_template = []
    if template_file:
        tpl_path = os.path.join(work_dir, "template.docx")
        with open(tpl_path, "wb") as f: f.write(template_file.getvalue())
        cmd_template = [f"--reference-doc={tpl_path}"]

    # 构建命令
    output_docx = os.path.join(work_dir, "preview_output.docx")
    cmd = [
        "pandoc", md_path,
        f"--metadata-file={yaml_path}",
        "--filter", CROSSREF_CMD,
        "--resource-path=.",
        "-o", output_docx
    ]
    if opt_toc: cmd.append("--toc")
    if opt_num: cmd.append("--number-sections")
    cmd.extend(cmd_template)

    res = subprocess.run(cmd, cwd=work_dir, capture_output=True, text=True)
    
    if res.returncode == 0:
        return output_docx, None
    else:
        return None, res.stderr

def convert_docx_to_pdf(docx_path, work_dir):
    """LibreOffice 核心转换逻辑：DOCX -> PDF"""
    if not HAS_LO:
        return None, "服务器未安装 LibreOffice，无法预览 PDF。请检查 packages.txt。"
    
    # 命令行调用 LibreOffice 转 PDF
    # --headless: 不启动图形界面
    # --convert-to pdf: 转换格式
    # --outdir: 输出目录
    cmd = [
        LO_CMD, "--headless", "--convert-to", "pdf", 
        docx_path, "--outdir", work_dir
    ]
    
    try:
        res = subprocess.run(cmd, cwd=work_dir, capture_output=True, text=True)
        # LibreOffice 成功时通常不会有报错，输出文件名同名，后缀改为 pdf
        pdf_filename = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
        pdf_path = os.path.join(work_dir, pdf_filename)
        
        if os.path.exists(pdf_path):
            return pdf_path, None
        else:
            return None, f"PDF 生成失败: {res.stderr}"
    except Exception as e:
        return None, str(e)

# ==========================================
# 🎨 3. 界面逻辑
# ==========================================

st.set_page_config(page_title="Pandoc Pro PDF", layout="wide", page_icon="📑")
st.title("Pandoc Pro: 真实 Word/PDF 预览")

# Session State
if 'pdf_base64' not in st.session_state: st.session_state['pdf_base64'] = None
if 'docx_data' not in st.session_state: st.session_state['docx_data'] = None

DEFAULT_YAML = """---
lang: zh-CN
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

with st.sidebar:
    st.header("1. 上传 Zip")
    source_file = st.file_uploader("文件上传", type=["zip", "md"])
    
    st.header("2. 样式模板")
    template_file = st.file_uploader("templates.docx", type=["docx"])
    
    st.header("3. 选项")
    opt_toc = st.checkbox("生成目录", False)
    opt_num = st.checkbox("章节编号", True)
    output_name = st.text_input("输出文件名", "paper_final")
    
    st.divider()
    if not HAS_LO:
        st.error("⚠️ 未检测到 LibreOffice，PDF 预览功能将不可用。请确保已添加 packages.txt。")

# Tabs
tab1, tab2 = st.tabs(["👁️ Word转PDF 真实预览", "⚙️ 配置"])

with tab2:
    yaml_content = st.text_area("YAML 配置", DEFAULT_YAML, height=400)

with tab1:
    col1, col2 = st.columns([1, 3]) # 左窄右宽
    
    with col1:
        st.info("💡 这里的预览是先生成 Word，再转为 PDF 的结果。所见即所得。")
        if source_file:
            if st.button("🔄 生成/刷新 预览", type="primary", use_container_width=True):
                with st.spinner("正在 Pandoc 编译 Word -> LibreOffice 转 PDF (首次较慢)..."):
                    with tempfile.TemporaryDirectory() as temp_dir:
                        # 1. 解压找到 MD
                        md_path, work_dir, err = unpack_and_find_md(source_file, temp_dir)
                        if err:
                            st.error(err)
                        else:
                            # 2. 生成 DOCX (中间产物)
                            docx_path, err = convert_to_docx(md_path, work_dir, yaml_content, template_file, opt_toc, opt_num)
                            if err:
                                st.error(f"Pandoc 转换失败:\n{err}")
                            else:
                                # 保存 DOCX 数据供下载
                                with open(docx_path, "rb") as f:
                                    st.session_state['docx_data'] = f.read()

                                # 3. Word -> PDF (预览用)
                                pdf_path, err = convert_docx_to_pdf(docx_path, work_dir)
                                if err:
                                    st.error(f"PDF 转换失败:\n{err}")
                                else:
                                    # 读取 PDF 并转为 Base64 以便嵌入浏览器
                                    with open(pdf_path, "rb") as f:
                                        base64_pdf = base64.b64encode(f.read()).decode('utf-8')
                                        st.session_state['pdf_base64'] = base64_pdf
                                        st.toast("预览已更新！", icon="✅")

            st.divider()
            st.subheader("📥 下载")
            if st.session_state['docx_data']:
                full_name = output_name if output_name.endswith(".docx") else output_name + ".docx"
                st.download_button(
                    label=f"下载 Word ({full_name})",
                    data=st.session_state['docx_data'],
                    file_name=full_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            else:
                st.caption("请先点击上方生成按钮")

    with col2:
        if st.session_state['pdf_base64']:
            # 使用 iframe 嵌入 PDF，利用浏览器原生的 PDF 阅读器 (自带缩放、翻页)
            pdf_display = f'<iframe src="data:application/pdf;base64,{st.session_state["pdf_base64"]}" width="100%" height="1000px" type="application/pdf"></iframe>'
            st.markdown(pdf_display, unsafe_allow_html=True)
        else:
            st.markdown(
                """
                <div style="border: 2px dashed #ccc; height: 800px; display: flex; align-items: center; justify-content: center; color: #888;">
                    <h3>👈 请上传文件并点击生成预览</h3>
                </div>
                """, 
                unsafe_allow_html=True
            )
