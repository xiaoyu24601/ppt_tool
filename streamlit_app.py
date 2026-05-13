import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import os
import tempfile
import shutil
import uuid
from pathlib import Path

st.set_page_config(page_title="Gamma PPTX 水印去除工具", page_icon="🔓", layout="centered")

st.title("🔓 Gamma PPTX 处理工具")
st.caption("去除水印 · 修复布局 · 文件不上传服务器 · 完全免费")

# 注册命名空间
NSMAP = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}
for prefix, uri in NSMAP.items():
    ET.register_namespace(prefix, uri)

OUTPUT_DIR = Path(__file__).parent / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)


def remove_gamma_watermark(input_path, output_path):
    temp_dir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(input_path, 'r') as z:
            z.extractall(temp_dir)

        changes = {'watermarks_removed': 0, 'layouts_cleaned': 0}

        for root_dir, dirs, files in os.walk(temp_dir):
            for file in files:
                if not file.endswith('.xml'):
                    continue
                filepath = os.path.join(root_dir, file)
                try:
                    tree = ET.parse(filepath)
                    root = tree.getroot()
                    file_changed = False

                    for elem in root.iter():
                        if elem.text and ('Made with Gamma' in elem.text or
                                          'gamma.app' in elem.text.lower() or
                                          'gamma' in str(elem.text).lower()):
                            elem.text = ''
                            file_changed = True
                            changes['watermarks_removed'] += 1

                        for attr_name, attr_value in list(elem.attrib.items()):
                            if attr_value and ('gamma' in str(attr_value).lower()):
                                if 'gamma.app' in str(attr_value).lower() or \
                                   'Made with Gamma' in str(attr_value):
                                    del elem.attrib[attr_name]
                                    file_changed = True
                                    changes['watermarks_removed'] += 1

                    if file_changed:
                        tree.write(filepath, xml_declaration=True, encoding='UTF-8')
                        if 'slideLayout' in file or 'slideMaster' in file:
                            changes['layouts_cleaned'] += 1

                except ET.ParseError:
                    continue

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    filepath = os.path.join(root_dir, file)
                    arcname = os.path.relpath(filepath, temp_dir)
                    zout.write(filepath, arcname)

        return {'success': True, 'message': f'✅ 处理完成！移除了 {changes["watermarks_removed"]} 个水印标记，清理了 {changes["layouts_cleaned"]} 个布局'}
    except Exception as e:
        return {'success': False, 'message': f'❌ 处理失败: {str(e)}'}
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def fix_pptx_layout(input_path, output_path):
    temp_dir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(input_path, 'r') as z:
            z.extractall(temp_dir)

        fixed_count = 0
        for root_dir, dirs, files in os.walk(temp_dir):
            for file in files:
                if not file.endswith('.xml'):
                    continue
                filepath = os.path.join(root_dir, file)
                try:
                    tree = ET.parse(filepath)
                    root = tree.getroot()
                    file_changed = False

                    for elem in root.iter():
                        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
                        if tag in ('rPr', 'defRPr'):
                            for child in elem:
                                child_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                                if child_tag == 'latin' and child.get('typeface', '') == '':
                                    child.set('typeface', 'Arial')
                                    file_changed = True
                                elif child_tag == 'ea' and child.get('typeface', '') == '':
                                    child.set('typeface', '微软雅黑')
                                    file_changed = True

                    if file_changed:
                        tree.write(filepath, xml_declaration=True, encoding='UTF-8')
                        fixed_count += 1

                except ET.ParseError:
                    continue

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    filepath = os.path.join(root_dir, file)
                    arcname = os.path.relpath(filepath, temp_dir)
                    zout.write(filepath, arcname)

        return {'success': True, 'message': f'✅ 修复完成！处理了 {fixed_count} 个文件'}
    except Exception as e:
        return {'success': False, 'message': f'❌ 修复失败: {str(e)}'}
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


# UI
col1, col2 = st.columns(2)
with col1:
    st.info("🔓 **去除水印**\n去除 \"Made with Gamma\" 标记")
with col2:
    st.info("🔧 **修复布局**\n修复空字体、对齐问题")

uploaded_file = st.file_uploader("上传 PPTX 文件", type=["pptx"], help="支持从 Gamma 导出的 .pptx 文件")

if uploaded_file:
    st.success(f"已上传: {uploaded_file.name}")

    c1, c2 = st.columns(2)

    with c1:
        if st.button("🔓 去除水印", use_container_width=True):
            with st.spinner("处理中..."):
                input_path = OUTPUT_DIR / f"input_{uuid.uuid4().hex[:8]}.pptx"
                output_path = OUTPUT_DIR / f"clean_{uuid.uuid4().hex[:8]}.pptx"

                with open(input_path, 'wb') as f:
                    f.write(uploaded_file.getbuffer())

                result = remove_gamma_watermark(str(input_path), str(output_path))

                if result['success']:
                    with open(output_path, 'rb') as f:
                        st.download_button(
                            label="📥 下载无水印版本",
                            data=f.read(),
                            file_name=f"去水印_{uploaded_file.name}",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                    st.success(result['message'])
                else:
                    st.error(result['message'])

                try:
                    os.unlink(input_path)
                    os.unlink(output_path)
                except:
                    pass

    with c2:
        if st.button("🔧 修复布局", use_container_width=True):
            with st.spinner("处理中..."):
                input_path = OUTPUT_DIR / f"input_{uuid.uuid4().hex[:8]}.pptx"
                output_path = OUTPUT_DIR / f"fixed_{uuid.uuid4().hex[:8]}.pptx"

                with open(input_path, 'wb') as f:
                    f.write(uploaded_file.getbuffer())

                result = fix_pptx_layout(str(input_path), str(output_path))

                if result['success']:
                    with open(output_path, 'rb') as f:
                        st.download_button(
                            label="📥 下载修复版本",
                            data=f.read(),
                            file_name=f"修复_{uploaded_file.name}",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                    st.success(result['message'])
                else:
                    st.error(result['message'])

                try:
                    os.unlink(input_path)
                    os.unlink(output_path)
                except:
                    pass
