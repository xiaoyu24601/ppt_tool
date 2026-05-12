"""
Gamma PPTX 处理工具 - 零依赖，纯标准库
功能：
1. 去除 Gamma 水印 "Made with Gamma"
2. 修复常见布局问题
3. 简单网页界面
"""
import zipfile
import xml.etree.ElementTree as ET
import os
import re
import shutil
import tempfile
from pathlib import Path
from http.server import HTTPServer, SimpleHTTPRequestHandler
import urllib.parse
import io

# Gamma 水印相关的命名空间
NSMAP = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}

# 注册命名空间
for prefix, uri in NSMAP.items():
    ET.register_namespace(prefix, uri)


def remove_gamma_watermark_from_pptx(input_path, output_path):
    """从 PPTX 文件中移除 Gamma 水印"""
    temp_dir = tempfile.mkdtemp()

    try:
        # 解压 PPTX (本质是 zip)
        with zipfile.ZipFile(input_path, 'r') as z:
            z.extractall(temp_dir)

        modified = False
        changes = {'watermarks_removed': 0, 'layouts_cleaned': 0}

        # 处理所有 XML 文件
        for root_dir, dirs, files in os.walk(temp_dir):
            for file in files:
                if not file.endswith('.xml'):
                    continue

                filepath = os.path.join(root_dir, file)
                try:
                    tree = ET.parse(filepath)
                    root = tree.getroot()
                    file_changed = False

                    # 搜索包含 "Made with Gamma" 或 "gamma" 的元素
                    for elem in root.iter():
                        # 检查文本内容
                        if elem.text and ('Made with Gamma' in elem.text or
                                          'gamma.app' in elem.text.lower() or
                                          'gamma' in str(elem.text).lower()):
                            elem.text = ''
                            file_changed = True
                            changes['watermarks_removed'] += 1

                        # 检查属性
                        for attr_name, attr_value in list(elem.attrib.items()):
                            if attr_value and ('gamma' in str(attr_value).lower()):
                                if 'gamma.app' in str(attr_value).lower() or \
                                   'Made with Gamma' in str(attr_value):
                                    del elem.attrib[attr_name]
                                    file_changed = True
                                    changes['watermarks_removed'] += 1

                    if file_changed:
                        tree.write(filepath, xml_declaration=True, encoding='UTF-8')
                        modified = True

                        if 'slideLayout' in file or 'slideMaster' in file:
                            changes['layouts_cleaned'] += 1

                except ET.ParseError:
                    continue

        # 重新打包为 PPTX
        if modified:
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                for root_dir, dirs, files in os.walk(temp_dir):
                    for file in files:
                        filepath = os.path.join(root_dir, file)
                        arcname = os.path.relpath(filepath, temp_dir)
                        zout.write(filepath, arcname)

            return {
                'success': True,
                'has_watermark': True,
                'message': f'处理完成！移除了 {changes["watermarks_removed"]} 个水印标记，清理了 {changes["layouts_cleaned"]} 个布局。',
                'changes': changes
            }
        else:
            # 没找到水印，直接复制原文件
            shutil.copy2(input_path, output_path)
            return {
                'success': True,
                'has_watermark': False,
                'message': '未检测到 Gamma 水印。',
                'changes': changes
            }

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def fix_pptx_layout(input_path, output_path):
    """修复 PPTX 常见布局问题"""
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

                    # 修复常见问题
                    for elem in root.iter():
                        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag

                        # 修复空字体引用
                        if tag in ('rPr', 'defRPr'):
                            for child in elem:
                                child_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                                if child_tag == 'latin' and child.get('typeface', '') == '':
                                    child.set('typeface', 'Arial')
                                    file_changed = True
                                elif child_tag == 'ea' and child.get('typeface', '') == '':
                                    child.set('typeface', '微软雅黑')
                                    file_changed = True

                        # 修复空文本框
                        if tag == 'sp' or tag == 'txBody':
                            has_text = False
                            for t_elem in elem.iter():
                                t_tag = t_elem.tag.split('}')[-1] if '}' in t_elem.tag else t_elem.tag
                                if t_tag == 't' and t_elem.text and t_elem.text.strip():
                                    has_text = True
                                    break

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

        return {
            'success': True,
            'message': f'修复完成！处理了 {fixed_count} 个文件。'
        }

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


# ================ 网页界面 ================

HTML_PAGE = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Gamma PPTX 处理工具</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh; display: flex; align-items: center; justify-content: center;
}
.container {
    background: white; border-radius: 20px; padding: 40px;
    max-width: 600px; width: 90%; box-shadow: 0 20px 60px rgba(0,0,0,0.3);
}
h1 { text-align: center; color: #333; margin-bottom: 10px; font-size: 24px; }
.subtitle { text-align: center; color: #888; margin-bottom: 30px; font-size: 14px; }
.upload-area {
    border: 2px dashed #ddd; border-radius: 12px; padding: 40px;
    text-align: center; cursor: pointer; transition: all 0.3s;
    margin-bottom: 20px;
}
.upload-area:hover { border-color: #667eea; background: #f8f9ff; }
.upload-area.dragover { border-color: #667eea; background: #eef0ff; }
.upload-icon { font-size: 48px; margin-bottom: 10px; }
input[type="file"] { display: none; }
.btn {
    display: block; width: 100%; padding: 14px; border: none; border-radius: 10px;
    font-size: 16px; cursor: pointer; margin-bottom: 10px; transition: all 0.3s;
    font-weight: 600;
}
.btn-remove {
    background: linear-gradient(135deg, #667eea, #764ba2); color: white;
}
.btn-fix {
    background: linear-gradient(135deg, #f093fb, #f5576c); color: white;
}
.btn:hover { transform: translateY(-2px); box-shadow: 0 4px 15px rgba(0,0,0,0.2); }
.btn:disabled { opacity: 0.5; cursor: not-allowed; transform: none; }
.message {
    padding: 15px; border-radius: 10px; margin-top: 20px; text-align: center;
    display: none;
}
.message.success { background: #d4edda; color: #155724; display: block; }
.message.error { background: #f8d7da; color: #721c24; display: block; }
.message.info { background: #d1ecf1; color: #0c5460; display: block; }
.download-btn {
    display: inline-block; margin-top: 10px; padding: 10px 20px;
    background: #28a745; color: white; text-decoration: none; border-radius: 8px;
}
.file-name { margin-top: 10px; color: #555; font-size: 14px; }
.features {
    display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 20px;
}
.feature {
    background: #f8f9fa; padding: 12px; border-radius: 8px; text-align: center;
    font-size: 13px; color: #555;
}
.feature .icon { font-size: 20px; }
</style>
</head>
<body>
<div class="container">
    <h1>Gamma PPTX 处理工具</h1>
    <p class="subtitle">去除水印 · 修复布局 · 完全免费 · 本地处理</p>

    <div class="features">
        <div class="feature">
            <div class="icon">🔓</div>
            去除 "Made with Gamma" 水印
        </div>
        <div class="feature">
            <div class="icon">🔧</div>
            修复导出布局问题
        </div>
        <div class="feature">
            <div class="icon">🔒</div>
            文件不上传，本地处理
        </div>
        <div class="feature">
            <div class="icon">📥</div>
            完美兼容 WPS/Office
        </div>
    </div>

    <div class="upload-area" id="uploadArea" onclick="document.getElementById('fileInput').click()">
        <div class="upload-icon">📁</div>
        <p>点击选择或拖拽 PPTX 文件到这里</p>
        <p style="font-size:12px;color:#aaa;margin-top:5px;">支持 .pptx 格式</p>
    </div>

    <input type="file" id="fileInput" accept=".pptx">
    <p class="file-name" id="fileName"></p>

    <button class="btn btn-remove" id="btnRemove" disabled onclick="processFile('remove')">
        去除水印
    </button>
    <button class="btn btn-fix" id="btnFix" disabled onclick="processFile('fix')">
        修复布局
    </button>

    <div id="message"></div>
    <div id="loading" style="display:none; text-align:center; padding:20px;">
        处理中... ⏳
    </div>
</div>

<script>
let selectedFile = null;

document.getElementById('fileInput').addEventListener('change', function(e) {
    selectedFile = e.target.files[0];
    updateUI();
    showMessage('');
});

const uploadArea = document.getElementById('uploadArea');
uploadArea.addEventListener('dragover', (e) => { e.preventDefault(); uploadArea.classList.add('dragover'); });
uploadArea.addEventListener('dragleave', () => { uploadArea.classList.remove('dragover'); });
uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    selectedFile = e.dataTransfer.files[0];
    updateUI();
});

function updateUI() {
    if (selectedFile && selectedFile.name.endsWith('.pptx')) {
        document.getElementById('fileName').textContent = '已选择: ' + selectedFile.name;
        document.getElementById('btnRemove').disabled = false;
        document.getElementById('btnFix').disabled = false;
    } else {
        document.getElementById('btnRemove').disabled = true;
        document.getElementById('btnFix').disabled = true;
    }
}

function showMessage(text, type) {
    const msg = document.getElementById('message');
    msg.textContent = text;
    msg.className = 'message ' + (type || '');
}

async function processFile(action) {
    if (!selectedFile) return;

    document.getElementById('loading').style.display = 'block';
    document.getElementById('btnRemove').disabled = true;
    document.getElementById('btnFix').disabled = true;
    showMessage('');

    const formData = new FormData();
    formData.append('file', selectedFile);
    formData.append('action', action);

    try {
        const resp = await fetch('/process', { method: 'POST', body: formData });
        const result = await resp.json();

        document.getElementById('loading').style.display = 'none';

        if (result.success) {
            let html = result.message;
            if (result.download_url) {
                html += '<br><a class="download-btn" href="' + result.download_url + '" download>下载处理后的文件</a>';
            }
            showMessage(html, 'success');
        } else {
            showMessage(result.error || '处理失败', 'error');
            document.getElementById('btnRemove').disabled = false;
            document.getElementById('btnFix').disabled = false;
        }
    } catch (err) {
        document.getElementById('loading').style.display = 'none';
        showMessage('网络错误: ' + err.message, 'error');
        document.getElementById('btnRemove').disabled = false;
        document.getElementById('btnFix').disabled = false;
    }
}
</script>
</body>
</html>
"""

import json
import uuid

OUTPUT_DIR = Path(__file__).parent / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)


class PPTXHandler(SimpleHTTPRequestHandler):
    """处理 PPTX 上传和下载"""

    def do_GET(self):
        if self.path == '/' or self.path == '/index.html':
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(HTML_PAGE.encode('utf-8'))
        elif self.path.startswith('/outputs/'):
            # 下载处理后的文件
            filename = os.path.basename(self.path)
            filepath = OUTPUT_DIR / filename
            if filepath.exists():
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
                self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
                self.end_headers()
                with open(filepath, 'rb') as f:
                    self.wfile.write(f.read())
            else:
                self.send_error(404)
        else:
            self.send_error(404)

    def do_POST(self):
        if self.path == '/process':
            content_type = self.headers.get('Content-Type', '')
            if 'multipart/form-data' not in content_type:
                self.send_error(400)
                return

            # 解析 multipart form data
            boundary = content_type.split('boundary=')[1].encode()
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)

            # 简单解析 multipart
            parts = body.split(b'--' + boundary)
            file_data = None
            action = 'remove'

            for part in parts:
                if b'Content-Disposition' not in part:
                    continue

                headers, _, content = part.partition(b'\r\n\r\n')
                content = content.rstrip(b'\r\n--')

                if b'name="file"' in headers:
                    # 找文件内容
                    file_data = content
                elif b'name="action"' in headers:
                    action = content.decode('utf-8').strip()

            if not file_data:
                self.send_json({'success': False, 'error': '未收到文件'})
                return

            # 保存上传的文件
            input_filename = f"input_{uuid.uuid4().hex[:8]}.pptx"
            input_path = OUTPUT_DIR / input_filename

            # 去掉 multipart 头部
            # 找文件内容的开始（跳过 Content-Type 行）
            file_start = file_data.find(b'\r\n\r\n')
            if file_start != -1:
                file_data = file_data[file_start + 4:]

            with open(input_path, 'wb') as f:
                f.write(file_data)

            # 处理文件
            output_filename = f"processed_{uuid.uuid4().hex[:8]}.pptx"
            output_path = OUTPUT_DIR / output_filename

            try:
                if action == 'remove':
                    result = remove_gamma_watermark_from_pptx(str(input_path), str(output_path))
                else:
                    result = fix_pptx_layout(str(input_path), str(output_path))

                if result['success']:
                    result['download_url'] = f'/outputs/{output_filename}'

                self.send_json(result)
            except Exception as e:
                self.send_json({'success': False, 'error': str(e)})
            finally:
                # 清理输入文件
                try:
                    os.unlink(input_path)
                except:
                    pass
        else:
            self.send_error(404)

    def send_json(self, data):
        self.send_response(200)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))


def main():
    port = 8999
    server = HTTPServer(('localhost', port), PPTXHandler)
    print(f'工具已启动: http://localhost:{port}')
    print('按 Ctrl+C 停止')
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\n已停止')
        server.shutdown()


if __name__ == '__main__':
    main()
