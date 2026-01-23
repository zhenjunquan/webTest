import streamlit as st
import subprocess
import os
import tarfile
import urllib.request
import tempfile
import sys
import zipfile
import streamlit.components.v1 as components
from docx import Document
from docx.shared import Pt, RGBColor

# ==========================================
# 🛠️ 1. 环境配置
# ==========================================
def install_linux_tools():
    base_dir = os.getcwd()
    bin_dir = os.path.join(base_dir, "bin")
    pandoc_exe = os.path.join(bin_dir, "pandoc")
    crossref_exe = os.path.join(bin_dir, "pandoc-crossref")
    if os.path.exists(pandoc_exe) and os.path.exists(crossref_exe): return bin_dir
    if not os.path.exists(bin_dir): os.makedirs(bin_dir)
    
    # 下载 Pandoc
    PANDOC_VER = "3.1.12.3"
    p_url = f"https://github.com/jgm/pandoc/releases/download/{PANDOC_VER}/pandoc-{PANDOC_VER}-linux-amd64.tar.gz"
    try:
        t_path, _ = urllib.request.urlretrieve(p_url)
        with tarfile.open(t_path) as t:
            for m in t.getmembers():
                if m.name.endswith("bin/pandoc"):
                    m.name = "pandoc"; t.extract(m, bin_dir)
    except: pass
    
    # 下载 Crossref
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
# 🤖 2. AI 核心
# ==========================================
import openai

def ask_ai_for_yaml(api_key, base_url, user_req):
    client = openai.OpenAI(api_key=api_key, base_url=base_url)
    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": f"生成Pandoc meta.yaml内容，要求：{user_req}。只返回YAML。"}]
        )
        return response.choices[0].message.content
    except Exception as e: return f"Error: {e}"

def ask_ai_for_template_code(api_key, base_url, user_req):
    client = openai.OpenAI(api_key=api_key, base_url=base_url)
    prompt = f"写Python代码使用python-docx生成generated_template.docx，样式要求：{user_req}。只返回代码。"
    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}]
        )
        return response.choices[0].message.content
    except: return None

# ==========================================
# 🎨 3. 样式定义
# ==========================================
A4_CSS = """
<style>
    body { background-color: #525659; padding: 40px 0; display: flex; justify-content: center; margin: 0; }
    .markdown-body {
        width: 21cm; min-height: 29.7cm; padding: 2cm;
        background: white; color: black;
        font-family: "Times New Roman", "SimSun", serif;
        box-shadow: 0 0 15px rgba(0,0,0,0.5);
    }
    img { max-width: 100%; height: auto; }
    table { border-collapse: collapse; width: 100%; margin: 1em 0; }
    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
    th { background-color: #f2f2f2; }
</style>
"""

ZOOM_SCRIPT = """
<div style="position:fixed;top:20px;right:30px;background:rgba(0,0,0,0.8);padding:8px;border-radius:20px;color:white;z-index:9999;">
    <button onclick="z(-0.1)" style="background:none;border:none;color:white;font-size:18px;cursor:pointer">－</button>
    <span id="zv">100%</span>
    <button onclick="z(0.1)" style="background:none;border:none;color:white;font-size:18px;cursor:pointer">＋</button>
</div>
<script>
let c=1.0;
function z(d){
    c+=d; if(c<0.3)c=0.3;
    let b=document.querySelector('.markdown-body');
    if(b){ b.style.transform=`scale(${c})`; b.style.transformOrigin='top center'; b.style.marginBottom=`${(c-1)*29.7}cm`; }
    document.getElementById('zv').innerText=Math.round(c*100)+"%";
}
</script>
"""

# ==========================================
# 🚀 4. 主程序
# ==========================================
st.set_page_config(page_title="Pandoc Pro", layout="wide", page_icon="📄")

# --- Session State 初始化 (关键修复点) ---
if 'view_mode' not in st.session_state: st.session_state['view_mode'] = 'setup'
if 'preview_html' not in st.session_state: st.session_state['preview_html'] = None
if 'docx_data' not in st.session_state: st.session_state['docx_data'] = None
if 'yaml_content' not in st.session_state: st.session_state['yaml_content'] = ""
# 🔴 修复：初始化文件名，防止 NameError
if 'output_filename' not in st.session_state: st.session_state['output_filename'] = "paper_final"

# ----------------- 视图 1: 配置页 -----------------
if st.session_state['view_mode'] == 'setup':
    st.title("📄 Pandoc 文档工场 (修复版)")
    
    with st.expander("🤖 AI 辅助配置 (可选)", expanded=False):
        api_key = st.text_input("OpenAI Key", type="password")
        base_url = st.text_input("Base URL", value="https://api.openai.com/v1")
        req = st.text_input("描述 YAML 配置")
        if st.button("生成 YAML") and api_key:
            res = ask_ai_for_yaml(api_key, base_url, req)
            if "Error" not in res: st.session_state['yaml_content'] = res.replace("```yaml","").replace("```","").strip()

    col1, col2 = st.columns([1, 1.2])
    
    with col1:
        source_file = st.file_uploader("📂 上传 Zip/MD", type=["zip", "md"])
        template_file = st.file_uploader("🎨 样式模板 (.docx)", type=["docx"])
        
        # 🔴 修复：使用 key 直接绑定 session_state，这样输入的值会被自动记住
        st.text_input("输出文件名", key="output_filename") 
    
    with col2:
        if not st.session_state['yaml_content']:
            st.session_state['yaml_content'] = "---\nlang: en\neqnos: true\n---"
        yaml_content = st.text_area("Meta.yaml", st.session_state['yaml_content'], height=300)

    st.divider()
    if st.button("🚀 生成 Word 并预览", type="primary", use_container_width=True):
        if source_file:
            with st.spinner("1. 生成 Word -> 2. 解析 Word 为 HTML..."):
                with tempfile.TemporaryDirectory() as temp_dir:
                    # 解压
                    if source_file.name.endswith('.zip'):
                        zip_path = os.path.join(temp_dir, "upload.zip")
                        with open(zip_path, "wb") as f: f.write(source_file.getvalue())
                        with zipfile.ZipFile(zip_path, 'r') as z: z.extractall(temp_dir)
                    else:
                        with open(os.path.join(temp_dir, source_file.name), "wb") as f: f.write(source_file.getvalue())
                    
                    # 找 MD
                    md_path = None
                    work_dir = temp_dir
                    for r, d, f in os.walk(temp_dir):
                        for file in f:
                            if file.endswith(".md"): 
                                md_path = os.path.join(r, file)
                                work_dir = r
                                break
                    
                    if md_path:
                        # 写入 YAML
                        with open(os.path.join(work_dir, "meta.yaml"), "w", encoding="utf-8") as f: f.write(yaml_content)
                        
                        # 写入模板
                        cmd_opts = []
                        if template_file:
                            tpl_path = os.path.join(work_dir, "template.docx")
                            with open(tpl_path, "wb") as f: f.write(template_file.getvalue())
                            cmd_opts = [f"--reference-doc={tpl_path}"]

                        # 生成 DOCX
                        docx_out = os.path.join(work_dir, "output.docx")
                        cmd_build = [
                            "pandoc", md_path, 
                            "--metadata-file=meta.yaml", 
                            "--filter", CROSSREF_CMD,
                            "--resource-path=.",
                            "-o", docx_out
                        ] + cmd_opts
                        
                        subprocess.run(cmd_build, cwd=work_dir)
                        
                        if os.path.exists(docx_out):
                            with open(docx_out, "rb") as f: st.session_state['docx_data'] = f.read()
                            
                            # DOCX -> HTML (逆向预览)
                            cmd_preview = [
                                "pandoc", docx_out, 
                                "-f", "docx",
                                "-t", "html5",
                                "--embed-resources", 
                                "--standalone",
                                "--mathjax"
                            ]
                            
                            res_html = subprocess.run(cmd_preview, cwd=work_dir, capture_output=True, text=True)
                            
                            if res_html.returncode == 0:
                                final_html = A4_CSS + '<div class="markdown-body">' + res_html.stdout + '</div>' + ZOOM_SCRIPT
                                st.session_state['preview_html'] = final_html
                                st.session_state['view_mode'] = 'preview'
                                st.rerun()
                            else:
                                st.error(f"生成预览失败: {res_html.stderr}")
                        else:
                            st.error("Word 生成失败")
        else:
            st.error("请先上传文件")

# ----------------- 视图 2: 预览页 -----------------
elif st.session_state['view_mode'] == 'preview':
    with st.sidebar:
        if st.button("⬅️ 返回修改"):
            st.session_state['view_mode'] = 'setup'
            st.rerun()
        st.divider()
        if st.session_state['docx_data']:
            # 🔴 修复：从 session_state 获取文件名
            name = st.session_state.get('output_filename', 'paper_final')
            fn = name if name.endswith(".docx") else name+".docx"
            
            st.download_button("📥 下载 Word 文档", st.session_state['docx_data'], fn, type="primary")
            st.info("提示：此预览是由生成的 Word 文档反向转换而来，所见即所得。")

    st.markdown("""<style>.block-container { padding-top: 1rem; padding-bottom: 0rem; } header { visibility: hidden; }</style>""", unsafe_allow_html=True)
    if st.session_state['preview_html']:
        components.html(st.session_state['preview_html'], height=1200, scrolling=True)
