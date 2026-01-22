import streamlit as st
import subprocess
import os
import tarfile
import urllib.request
import tempfile
import sys
import zipfile
import shutil

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

if sys.platform.startswith("linux"):
    local_bin = install_linux_tools()
    os.environ["PATH"] = local_bin + os.pathsep + os.environ["PATH"]
    CROSSREF_CMD = os.path.join(local_bin, "pandoc-crossref")
else:
    CROSSREF_CMD = "pandoc-crossref"

# ==========================================
# 🎨 2. 界面与核心逻辑
# ==========================================

st.set_page_config(page_title="Pandoc Pro", layout="wide", page_icon="📑")
st.title("📑 Markdown 转 Word (修复版)")

# 初始化 Session State (解决下载没反应的问题)
if 'convert_success' not in st.session_state:
    st.session_state['convert_success'] = False
    st.session_state['docx_data'] = None
    st.session_state['log_info'] = ""

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

with st.sidebar:
    st.header("📂 1. 文件上传")
    st.info("💡 请上传 Zip 包，包含 .md 文件和所有图片文件夹。")
    source_file = st.file_uploader("上传 Zip 文件", type=["zip"])
    
    st.header("🎨 2. 样式 & 选项")
    template_file = st.file_uploader("样式模板 (templates.docx)", type=["docx"])
    
    opt_toc = st.checkbox("生成目录 (--toc)", False)
    opt_num = st.checkbox("章节编号 (--number-sections)", True)
    output_name = st.text_input("输出文件名", "paper_final")

# --- 主逻辑 ---
tab1, tab2 = st.tabs(["🚀 转换 & 下载", "⚙️ 配置 (YAML)"])

with tab2:
    yaml_content = st.text_area("编辑 YAML", DEFAULT_YAML, height=400)

with tab1:
    if source_file:
        # Step 1: 转换按钮
        if st.button("🔄 开始转换 (第一步)", type="primary"):
            # 清除旧状态
            st.session_state['convert_success'] = False
            st.session_state['docx_data'] = None
            
            with st.spinner("正在解压并编译..."):
                with tempfile.TemporaryDirectory() as temp_dir:
                    # 1. 解压 Zip
                    zip_path = os.path.join(temp_dir, "upload.zip")
                    with open(zip_path, "wb") as f:
                        f.write(source_file.getvalue())
                    
                    try:
                        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                            zip_ref.extractall(temp_dir)
                    except Exception as e:
                        st.error(f"Zip 解压失败: {e}")
                        st.stop()

                    # 2. 深度搜索 .md 文件 (解决路径问题)
                    md_path = None
                    md_rel_dir = "" # MD 文件所在的文件夹
                    
                    file_structure = [] # 用于调试
                    for root, dirs, files in os.walk(temp_dir):
                        for file in files:
                            file_structure.append(os.path.join(root, file))
                            if file.endswith(".md") and not file.startswith("__"):
                                md_path = os.path.join(root, file)
                                # 关键：记录 MD 文件所在的目录
                                md_rel_dir = root 
                                break
                        if md_path: break
                    
                    if not md_path:
                        st.error("❌ Zip 包里没找到 .md 文件！")
                        st.stop()

                    # 3. 准备资源路径
                    # 告诉 Pandoc 在 MD 文件所在的目录找图片
                    # 并显式添加 --resource-path
                    
                    # 保存 YAML
                    yaml_path = os.path.join(md_rel_dir, "meta.yaml")
                    with open(yaml_path, "w", encoding="utf-8") as f:
                        f.write(yaml_content)
                    
                    # 保存模板
                    cmd_template = []
                    if template_file:
                        tpl_path = os.path.join(md_rel_dir, "template.docx")
                        with open(tpl_path, "wb") as f:
                            f.write(template_file.getvalue())
                        cmd_template = [f"--reference-doc={tpl_path}"]

                    # 4. 构建命令
                    cmd = [
                        "pandoc", 
                        os.path.basename(md_path), # 只传文件名
                        f"--metadata-file=meta.yaml", 
                        "--filter", CROSSREF_CMD,
                        "--resource-path=.", # 强制在当前目录找图片
                        "-o", "output.docx"
                    ]
                    if opt_toc: cmd.append("--toc")
                    if opt_num: cmd.append("--number-sections")
                    cmd.extend(cmd_template)

                    # 5. 执行 (关键：cwd 设为 MD 文件所在的目录)
                    res = subprocess.run(cmd, cwd=md_rel_dir, capture_output=True, text=True)

                    if res.returncode == 0:
                        out_path = os.path.join(md_rel_dir, "output.docx")
                        with open(out_path, "rb") as f:
                            st.session_state['docx_data'] = f.read()
                        st.session_state['convert_success'] = True
                        
                        # 记录一些调试信息给用户看
                        msg = f"✅ 转换成功！\n\n📂 **工作目录**: `{md_rel_dir}`\n📄 **处理文件**: `{os.path.basename(md_path)}`"
                        # 检查图片文件夹是否存在
                        if os.path.exists(os.path.join(md_rel_dir, "images")):
                            msg += "\n🖼️ **检测**: 发现 `images` 文件夹，图片应该正常。"
                        elif os.path.exists(os.path.join(md_rel_dir, "assets")):
                            msg += "\n🖼️ **检测**: 发现 `assets` 文件夹，图片应该正常。"
                        else:
                            msg += "\n⚠️ **注意**: 未在 MD 同级目录发现 `images` 或 `assets` 文件夹。如果你的文档有图片，请检查 Zip 结构。"
                        
                        st.session_state['log_info'] = msg
                    else:
                        st.error("❌ 转换失败")
                        st.code(res.stderr)
                        st.warning("调试：Zip 包内的文件结构如下：")
                        st.json(file_structure)

        # Step 2: 下载按钮 (独立显示)
        if st.session_state['convert_success'] and st.session_state['docx_data']:
            st.success(st.session_state['log_info'])
            
            full_name = output_name if output_name.endswith(".docx") else output_name + ".docx"
            
            # 这里是解决“点击无反应”的关键：直接提供数据，不再运行逻辑
            st.download_button(
                label=f"📥 点击下载 Word 文档 ({full_name})",
                data=st.session_state['docx_data'],
                file_name=full_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary"
            )
    else:
        st.info("👈 请先上传 Zip 文件")
