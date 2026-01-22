import streamlit as st
import subprocess
import os
import tarfile
import urllib.request
import tempfile
import sys
import zipfile
import streamlit.components.v1 as components

# ==========================================
# 🛠️ 1. 环境配置 (保持不变)
# ==========================================
def install_linux_tools():
    """云端自动安装 Pandoc"""
    base_dir = os.getcwd()
    bin_dir = os.path.join(base_dir, "bin")
    pandoc_exe = os.path.join(bin_dir, "pandoc")
    crossref_exe = os.path.join(bin_dir, "pandoc-crossref")

    if os.path.exists(pandoc_exe) and os.path.exists(crossref_exe):
        return bin_dir

    st.toast("正在初始化 Pandoc...", icon="🚀")
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

if sys.platform.startswith("linux"):
    local_bin = install_linux_tools()
    os.environ["PATH"] = local_bin + os.pathsep + os.environ["PATH"]
    CROSSREF_CMD = os.path.join(local_bin, "pandoc-crossref")
else:
    CROSSREF_CMD = "pandoc-crossref"

# ==========================================
# 📂 2. 文件解压逻辑
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

# ==========================================
# 🎨 3. 页面样式定义 (A4 & 全屏优化)
# ==========================================

# A4 纸张 CSS
A4_CSS = """
<style>
    body {
        background-color: #525659; /* 深色背景，护眼且突出白纸 */
        display: flex;
        justify-content: center;
        padding: 40px 0;
        margin: 0;
    }
    .markdown-body {
        box-sizing: border-box;
        width: 21cm; /* A4 宽度 */
        min-height: 29.7cm; /* A4 高度 */
        margin: 0 auto;
        padding: 2.54cm; /* 标准页边距 */
        background-color: white;
        box-shadow: 0 0 10px rgba(0,0,0,0.5);
        color: #000;
        font-family: "Times New Roman", "SimSun", serif; /* 衬线体更像论文 */
    }
    /* 适配图片 */
    img { max-width: 100%; }
</style>
"""

# JS 缩放脚本
ZOOM_SCRIPT = """
<style>
  #float-toolbar {
    position: fixed; top: 20px; right: 30px; z-index: 10000;
    background: rgba(0,0,0,0.7); padding: 8px 15px;
    border-radius: 30px; display: flex; align-items: center; gap: 15px;
    color: white; backdrop-filter: blur(5px);
  }
  .zoom-btn {
    cursor: pointer; border: none; background: transparent; color: white;
    font-size: 18px; display: flex; align-items: center; font-weight: bold;
  }
  .zoom-btn:hover { color: #4CAF50; }
  #zoom-val { font-size: 14px; font-family: monospace; min-width: 45px; text-align: center; }
</style>

<div id="float-toolbar">
    <button class="zoom-btn" onclick="changeZoom(-0.1)">－</button>
    <span id="zoom-val">100%</span>
    <button class="zoom-btn" onclick="changeZoom(0.1)">＋</button>
</div>

<script>
    let currentZoom = 1.0;
    function changeZoom(delta) {
        currentZoom += delta;
        if (currentZoom < 0.3) currentZoom = 0.3;
        
        // 核心缩放逻辑
        const body = document.querySelector('.markdown-body');
        if(body) {
            body.style.transform = `scale(${currentZoom})`;
            body.style.transformOrigin = 'top center';
            // 动态调整底部留白，防止缩放后重叠
            body.style.marginBottom = `${(currentZoom - 1) * 29.7}cm`; 
        }
        document.getElementById('zoom-val').innerText = Math.round(currentZoom * 100) + "%";
    }
</script>
"""

# ==========================================
# 🚀 4. 主程序逻辑
# ==========================================

st.set_page_config(page_title="Pandoc Pro", layout="wide", page_icon="📝")

# --- Session State 初始化 ---
if 'view_mode' not in st.session_state: st.session_state['view_mode'] = 'setup' # setup 或 preview
if 'preview_html' not in st.session_state: st.session_state['preview_html'] = None
if 'docx_data' not in st.session_state: st.session_state['docx_data'] = None
if 'file_name' not in st.session_state: st.session_state['file_name'] = "paper_final"

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

# ==========================================
# 📺 视图 1：配置与上传页
# ==========================================
if st.session_state['view_mode'] == 'setup':
    st.title("Pandoc Pro: 文档编译平台")
    st.markdown("上传 Markdown/Zip，生成 A4 仿真预览与 Word 文档。")

    col_conf, col_yaml = st.columns([1, 1.5])

    with col_conf:
        st.subheader("1. 文件与模板")
        source_file = st.file_uploader("📂 上传 Zip (含MD和图片)", type=["zip", "md"])
        template_file = st.file_uploader("🎨 样式模板 (templates.docx)", type=["docx"])
        
        st.subheader("2. 输出设置")
        opt_toc = st.checkbox("生成目录 (--toc)", False)
        opt_num = st.checkbox("章节编号 (--number-sections)", True)
        output_name = st.text_input("输出文件名", st.session_state['file_name'])

    with col_yaml:
        st.subheader("3. 元数据配置")
        yaml_content = st.text_area("Meta.yaml", DEFAULT_YAML, height=450)

    # 底部大按钮
    st.divider()
    if st.button("🚀 生成预览 & 转换", type="primary", use_container_width=True):
        if not source_file:
            st.error("请先上传文件！")
        else:
            st.session_state['file_name'] = output_name # 记住文件名
            with st.spinner("正在启动 Pandoc 引擎..."):
                with tempfile.TemporaryDirectory() as temp_dir:
                    md_path, work_dir, err = unpack_and_find_md(source_file, temp_dir)
                    if err:
                        st.error(err)
                    else:
                        # 保存配置
                        yaml_path = os.path.join(work_dir, "meta.yaml")
                        with open(yaml_path, "w", encoding="utf-8") as f: f.write(yaml_content)
                        
                        # 1. 转换 Word (用于下载)
                        cmd_template = []
                        if template_file:
                            tpl_path = os.path.join(work_dir, "template.docx")
                            with open(tpl_path, "wb") as f: f.write(template_file.getvalue())
                            cmd_template = [f"--reference-doc={tpl_path}"]

                        output_docx = os.path.join(work_dir, "final.docx")
                        cmd_docx = [
                            "pandoc", md_path,
                            f"--metadata-file={yaml_path}",
                            "--filter", CROSSREF_CMD,
                            "--resource-path=.",
                            "-o", output_docx
                        ]
                        if opt_toc: cmd_docx.append("--toc")
                        if opt_num: cmd_docx.append("--number-sections")
                        cmd_docx.extend(cmd_template)
                        
                        subprocess.run(cmd_docx, cwd=work_dir)
                        if os.path.exists(output_docx):
                            with open(output_docx, "rb") as f:
                                st.session_state['docx_data'] = f.read()

                        # 2. 转换 HTML (用于全屏预览)
                        cmd_html = [
                            "pandoc", md_path,
                            f"--metadata-file={yaml_path}",
                            "--filter", CROSSREF_CMD,
                            "--to", "html5",
                            "--embed-resources",
                            "--standalone",
                            "--mathjax",
                            "--css", "https://cdn.jsdelivr.net/npm/github-markdown-css/github-markdown.min.css"
                        ]
                        if opt_toc: cmd_html.append("--toc")
                        if opt_num: cmd_html.append("--number-sections")

                        res_html = subprocess.run(cmd_html, cwd=work_dir, capture_output=True, text=True)
                        if res_html.returncode == 0:
                            # 拼接：A4 CSS + HTML + Zoom JS
                            st.session_state['preview_html'] = A4_CSS + res_html.stdout + ZOOM_SCRIPT
                            # 切换视图状态！
                            st.session_state['view_mode'] = 'preview'
                            st.rerun() # 强制刷新页面进入预览模式
                        else:
                            st.error(f"预览生成失败: {res_html.stderr}")

# ==========================================
# 🖥️ 视图 2：全屏预览页 (沉浸模式)
# ==========================================
elif st.session_state['view_mode'] == 'preview':
    
    # --- 侧边栏：操作区 ---
    with st.sidebar:
        st.header("操作栏")
        
        # 返回按钮
        if st.button("⬅️ 返回修改", use_container_width=True):
            st.session_state['view_mode'] = 'setup'
            st.rerun()
            
        st.divider()
        
        # 下载按钮
        if st.session_state['docx_data']:
            fname = st.session_state['file_name']
            if not fname.endswith(".docx"): fname += ".docx"
            
            st.download_button(
                label="📥 下载 Word 文档",
                data=st.session_state['docx_data'],
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )
        
        st.info("提示：右侧为 HTML 仿真预览，排版与 Word 可能略有差异，但内容与公式一致。")
        
        # 高度控制
        st.divider()
        iframe_height = st.slider("预览窗口高度", 800, 3000, 1200)

    # --- 主区域：全屏 HTML ---
    # 移除顶部的 padding，让预览更沉浸
    st.markdown("""
        <style>
               .block-container { padding-top: 1rem; padding-bottom: 0rem; }
               header { visibility: hidden; }
        </style>
        """, unsafe_allow_html=True)

    if st.session_state['preview_html']:
        components.html(st.session_state['preview_html'], height=iframe_height, scrolling=True)
    else:
        st.error("预览数据丢失，请返回重新生成。")
