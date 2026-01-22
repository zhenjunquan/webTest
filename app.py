import streamlit as st
import subprocess
import os
import tarfile
import urllib.request
import tempfile
import sys
import zipfile
import shutil
import streamlit.components.v1 as components

# ==========================================
# 🛠️ 1. 环境配置 (无需修改)
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

if sys.platform.startswith("linux"):
    local_bin = install_linux_tools()
    os.environ["PATH"] = local_bin + os.pathsep + os.environ["PATH"]
    CROSSREF_CMD = os.path.join(local_bin, "pandoc-crossref")
else:
    CROSSREF_CMD = "pandoc-crossref"

# ==========================================
# 📂 2. 文件处理函数
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

    md_path = None
    work_dir = temp_dir
    
    for root, dirs, files in os.walk(temp_dir):
        for file in files:
            if file.endswith(".md") and not file.startswith("__"):
                md_path = os.path.join(root, file)
                work_dir = root 
                return md_path, work_dir, None
    
    return None, None, "未找到 .md 文件"

# ==========================================
# 🎨 3. 界面逻辑
# ==========================================

st.set_page_config(page_title="Pandoc Pro", layout="wide", page_icon="👁️")
st.title("Pandoc Pro: 预览与转换")

if 'preview_html' not in st.session_state: st.session_state['preview_html'] = None
if 'md_content' not in st.session_state: st.session_state['md_content'] = ""
if 'docx_data' not in st.session_state: st.session_state['docx_data'] = None

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

# --- JS 注入：悬浮缩放控制器 ---
# 这段代码会被拼接到生成的 HTML 里，从而在预览框内部生成按钮
ZOOM_SCRIPT = """
<style>
  #float-toolbar {
    position: fixed; top: 10px; right: 20px; z-index: 9999;
    background: rgba(255,255,255,0.9); padding: 5px 10px;
    border: 1px solid #ccc; border-radius: 20px;
    box-shadow: 0 4px 10px rgba(0,0,0,0.1);
    font-family: sans-serif; display: flex; align-items: center; gap: 8px;
  }
  .zoom-btn {
    cursor: pointer; border: none; background: #eee;
    width: 24px; height: 24px; border-radius: 50%; font-size: 16px;
    display: flex; align-items: center; justify-content: center;
    transition: background 0.2s;
  }
  .zoom-btn:hover { background: #ddd; }
  #zoom-val { font-size: 12px; font-weight: bold; color: #555; width: 40px; text-align: center; }
</style>

<div id="float-toolbar">
    <button class="zoom-btn" onclick="changeZoom(-0.1)" title="缩小">➖</button>
    <span id="zoom-val">100%</span>
    <button class="zoom-btn" onclick="changeZoom(0.1)" title="放大">➕</button>
</div>

<script>
    let currentZoom = 1.0;
    function changeZoom(delta) {
        currentZoom += delta;
        if (currentZoom < 0.2) currentZoom = 0.2;
        // 兼容性缩放处理
        document.body.style.zoom = currentZoom;
        document.getElementById('zoom-val').innerText = Math.round(currentZoom * 100) + "%";
    }
</script>
"""

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
    st.write("🔧 **预览窗口设置**")
    # 🔴 新增功能：外部控制预览框高度
    iframe_height = st.slider("预览框高度 (px)", min_value=600, max_value=2000, value=1000, step=100)

tab1, tab2, tab3 = st.tabs(["👁️ 实时预览", "📥 转换下载", "⚙️ 配置"])

# --- Tab 3: Config ---
with tab3:
    yaml_content = st.text_area("YAML 配置", DEFAULT_YAML, height=400)

# --- Tab 1: Preview ---
with tab1:
    if source_file:
        if st.button("🔄 刷新预览", type="primary", use_container_width=True):
            with st.spinner("渲染中..."):
                with tempfile.TemporaryDirectory() as temp_dir:
                    md_path, work_dir, err = unpack_and_find_md(source_file, temp_dir)
                    
                    if err:
                        st.error(err)
                    else:
                        with open(md_path, 'r', encoding='utf-8') as f:
                            st.session_state['md_content'] = f.read()
                        
                        yaml_path = os.path.join(work_dir, "meta.yaml")
                        with open(yaml_path, "w", encoding="utf-8") as f:
                            f.write(yaml_content)

                        cmd = [
                            "pandoc", md_path,
                            f"--metadata-file={yaml_path}",
                            "--filter", CROSSREF_CMD,
                            "--to", "html5",
                            "--embed-resources", 
                            "--standalone",
                            "--mathjax",
                            "--css", "https://cdn.jsdelivr.net/npm/github-markdown-css/github-markdown.min.css"
                        ]
                        if opt_toc: cmd.append("--toc")
                        if opt_num: cmd.append("--number-sections")
                        
                        res = subprocess.run(cmd, cwd=work_dir, capture_output=True, text=True)
                        
                        if res.returncode == 0:
                            # 🔴 核心修改：将 JS 缩放代码注入到 HTML 尾部
                            st.session_state['preview_html'] = res.stdout + ZOOM_SCRIPT
                        else:
                            st.error("预览失败")
                            st.code(res.stderr)

        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("MD 源码")
            if st.session_state['md_content']:
                with st.container(height=iframe_height, border=True):
                    st.markdown(st.session_state['md_content'])
            else:
                st.info("👈 等待加载")

        with col2:
            st.subheader("HTML 仿真预览")
            if st.session_state['preview_html']:
                # 🔴 核心修改：使用 slider 的值动态控制高度
                components.html(st.session_state['preview_html'], height=iframe_height, scrolling=True)
            else:
                st.info("👈 等待加载")
    else:
        st.info("请先上传文件")

# --- Tab 2: Download ---
with tab2:
    st.write("### 生成 Word 文档")
    if source_file:
        if st.button("🚀 开始转换 Word", type="primary"):
            with st.spinner("正在生成 DOCX..."):
                with tempfile.TemporaryDirectory() as temp_dir:
                    md_path, work_dir, err = unpack_and_find_md(source_file, temp_dir)
                    
                    if not err:
                        yaml_path = os.path.join(work_dir, "meta.yaml")
                        with open(yaml_path, "w", encoding="utf-8") as f: f.write(yaml_content)
                        
                        cmd_template = []
                        if template_file:
                            tpl_path = os.path.join(work_dir, "template.docx")
                            with open(tpl_path, "wb") as f: f.write(template_file.getvalue())
                            cmd_template = [f"--reference-doc={tpl_path}"]

                        cmd = [
                            "pandoc", md_path,
                            f"--metadata-file={yaml_path}",
                            "--filter", CROSSREF_CMD,
                            "--resource-path=.", 
                            "-o", "output.docx"
                        ]
                        if opt_toc: cmd.append("--toc")
                        if opt_num: cmd.append("--number-sections")
                        cmd.extend(cmd_template)

                        res = subprocess.run(cmd, cwd=work_dir, capture_output=True, text=True)

                        if res.returncode == 0:
                            out_path = os.path.join(work_dir, "output.docx")
                            with open(out_path, "rb") as f:
                                st.session_state['docx_data'] = f.read()
                            st.success("转换成功！")
                        else:
                            st.error("转换失败")
                            st.code(res.stderr)

        if st.session_state['docx_data']:
            full_name = output_name if output_name.endswith(".docx") else output_name + ".docx"
            st.download_button(
                label=f"📥 下载 {full_name}",
                data=st.session_state['docx_data'],
                file_name=full_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary"
            )
    else:
        st.info("请先上传文件")
