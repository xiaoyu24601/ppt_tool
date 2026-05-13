"""
PPT 工具箱 — AI 生成 · 美化 · 去水印 · 修复
管理密码 + 激活码 + 免 API 直接使用
"""
import streamlit as st
import hashlib, hmac, secrets, json, io, os, zipfile, tempfile, shutil, uuid, re
from pathlib import Path
from datetime import datetime, timedelta
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import xml.etree.ElementTree as ET
import requests

st.set_page_config(page_title="PPT 工具箱", page_icon="🪄", layout="wide")

# ==================== 配置 ====================

# DeepSeek API Key（你的密钥，内置在代码中）
DEEPSEEK_API_KEY = "sk-7e6523558c3a457c93b60b2f87a9155c"

# 管理密码 — 改成你自己的
ADMIN_PASSWORD = "lmy2026ppt"

# 激活码密钥盐
ACTIVATION_SALT = "ppt-tool-salt-2026"

# 预生成的有效激活码（用 generate_code() 生成）
# 格式: {"code": "激活码", "used": False, "created": "日期", "note": "备注"}
VALID_CODES = [
    {"code": "PPT-VIP-0001", "used": False, "created": "2026-05-13", "note": "测试码"},
    {"code": "PPT-VIP-0002", "used": False, "created": "2026-05-13", "note": "测试码"},
    {"code": "PPT-VIP-0003", "used": False, "created": "2026-05-13", "note": "测试码"},
]

# ==================== 激活码逻辑 ====================

def generate_code(prefix="PPT", index=1):
    """生成新激活码"""
    raw = f"{ACTIVATION_SALT}-{prefix}-{index}-{secrets.token_hex(4)}"
    hash_val = hashlib.sha256(raw.encode()).hexdigest()[:8].upper()
    return f"{prefix}-{hash_val}"

def validate_code(code):
    """验证激活码是否有效且未使用"""
    for entry in VALID_CODES:
        if entry["code"] == code and not entry["used"]:
            return entry
    return None

def mark_code_used(code):
    """标记激活码已使用"""
    for entry in VALID_CODES:
        if entry["code"] == code:
            entry["used"] = True
            entry["used_at"] = datetime.now().isoformat()
            return True
    return False

# ==================== 会话状态 ====================

if "ai_enabled" not in st.session_state:
    st.session_state.ai_enabled = False
if "auth_type" not in st.session_state:
    st.session_state.auth_type = None  # "admin" or "activation"
if "code_used" not in st.session_state:
    st.session_state.code_used = None

# ==================== 主题配置 ====================

THEMES = {
    "专业商务": {"primary": "#2C3E50", "secondary": "#3498DB", "accent": "#2ECC71", "font": "Arial"},
    "学术答辩": {"primary": "#003366", "secondary": "#C0392B", "accent": "#F39C12", "font": "Times New Roman"},
    "创业路演": {"primary": "#E74C3C", "secondary": "#F39C12", "accent": "#9B59B6", "font": "Arial"},
    "科技互联网": {"primary": "#6C63FF", "secondary": "#00D4FF", "accent": "#FF6B6B", "font": "Helvetica"},
    "教育培训": {"primary": "#27AE60", "secondary": "#2ECC71", "accent": "#3498DB", "font": "Arial"},
    "简约极简": {"primary": "#333333", "secondary": "#666666", "accent": "#999999", "font": "Helvetica"},
    "清新自然": {"primary": "#16A085", "secondary": "#27AE60", "accent": "#F1C40F", "font": "Arial"},
    "暗夜模式": {"primary": "#FFFFFF", "secondary": "#3498DB", "accent": "#E74C3C", "font": "Arial", "dark": True},
    "创意设计": {"primary": "#9B59B6", "secondary": "#E91E63", "accent": "#FF9800", "font": "Comic Sans MS"},
    "政府汇报": {"primary": "#8B0000", "secondary": "#B8860B", "accent": "#006400", "font": "SimHei"},
}

OUTLINE_LIBRARY = {
    "专业商务": ["项目背景", "市场分析", "解决方案", "竞争优势", "执行计划", "预期成果"],
    "学术答辩": ["研究背景", "文献综述", "研究方法", "实验结果", "分析讨论", "结论展望"],
    "创业路演": ["痛点分析", "解决方案", "市场规模", "商业模式", "团队介绍", "融资计划"],
    "科技互联网": ["行业趋势", "技术架构", "产品功能", "用户场景", "数据指标", "未来路线"],
    "教育培训": ["课程目标", "核心概念", "案例讲解", "实操演练", "常见误区", "总结回顾"],
    "简约极简": ["背景", "问题", "方案", "优势", "计划", "总结"],
    "清新自然": ["引入", "现状", "对策", "案例", "成果", "展望"],
    "暗夜模式": ["背景", "挑战", "策略", "执行", "数据", "下一步"],
    "创意设计": ["灵感来源", "设计理念", "视觉方案", "细节展示", "应用场景", "迭代方向"],
    "政府汇报": ["工作概述", "重点任务", "进展成效", "问题分析", "改进措施", "下步计划"],
}

AI_SYSTEM_PROMPT = """你是一个专业的 PPT 内容策划专家。根据用户提供的描述生成结构化的 PPT 内容。

用户输入的可能是简短主题，也可能是详细描述。你需要根据内容长度和复杂度做出相应调整。

输出严格的 JSON 格式（不要 markdown 代码块）：
{
  "title": "PPT 总标题",
  "subtitle": "副标题",
  "slides": [
    {
      "title": "页面标题",
      "bullets": ["要点1", "要点2", "要点3"],
      "notes": "演讲备注"
    }
  ]
}

要求：
- 包含封面页（title + subtitle）和 4-8 页内容
- 每页 3-5 个要点，每个要点 15-40 字
- 标题精简有力（15 字以内）
- 如果用到了用户描述中的具体数据、案例，务必保留
- 内容结构清晰，逻辑递进"""

# ==================== 注册 XML 命名空间 ====================

for p, u in {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}.items():
    ET.register_namespace(p, u)

OUTPUT_DIR = Path(__file__).parent / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)


# ==================== AI PPT 生成 ====================

def call_deepseek(prompt, max_tokens=4096):
    """调用 DeepSeek API"""
    resp = requests.post(
        "https://api.deepseek.com/chat/completions",
        headers={"Authorization": f"Bearer {DEEPSEEK_API_KEY}", "Content-Type": "application/json"},
        json={
            "model": "deepseek-chat",
            "messages": [
                {"role": "system", "content": AI_SYSTEM_PROMPT},
                {"role": "user", "content": prompt},
            ],
            "temperature": 0.7,
            "max_tokens": max_tokens,
        },
        timeout=120,
    )
    if resp.status_code != 200:
        raise Exception(f"API 错误 ({resp.status_code}): {resp.text[:200]}")
    return resp.json()["choices"][0]["message"]["content"]


def parse_ai_response(raw):
    """解析 AI 返回的 JSON"""
    raw = raw.strip()
    if raw.startswith("```"):
        raw = re.sub(r'^```\w*\n', '', raw)
        raw = re.sub(r'\n```$', '', raw)
    return json.loads(raw)


def build_pptx(data, theme_name, slide_count):
    """用 python-pptx 构建 PPTX"""
    t = THEMES.get(theme_name, THEMES["专业商务"])
    is_dark = t.get("dark", False)
    pc, sc, ac = t["primary"], t["secondary"], t["accent"]
    bg = "#1E1E1E" if is_dark else "#FFFFFF"
    tc = "#FFFFFF" if is_dark else "#333333"

    def h(c):
        c = c.lstrip("#")
        return RGBColor(int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16))

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    title_text = data.get("title", "未命名")
    subtitle_text = data.get("subtitle", "")
    slides = data.get("slides", [])

    if not slides:
        # 如果 AI 没给 slides，自动拆分
        slides = [{"title": title_text, "bullets": ["请查看详细内容"], "notes": ""}]

    # === 封面 ===
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl.background.fill.solid(); sl.background.fill.fore_color.rgb = h(bg)

    # 装饰条
    bar = sl.shapes.add_shape(1, Inches(0), Inches(0), Inches(0.12), Inches(7.5))
    bar.fill.solid(); bar.fill.fore_color.rgb = h(sc); bar.line.fill.background()

    # 标题
    tb = sl.shapes.add_textbox(Inches(1.2), Inches(1.8), Inches(11), Inches(2.2))
    p = tb.text_frame.paragraphs[0]
    p.text = title_text[:60]; p.font.size = Pt(48); p.font.bold = True
    p.font.color.rgb = h(pc); p.font.name = t["font"]

    # 副标题
    if subtitle_text:
        tb2 = sl.shapes.add_textbox(Inches(1.2), Inches(4.2), Inches(11), Inches(1))
        p2 = tb2.text_frame.paragraphs[0]
        p2.text = subtitle_text[:100]; p2.font.size = Pt(20)
        p2.font.color.rgb = h("#999999"); p2.font.name = t["font"]

    # 日期
    tb3 = sl.shapes.add_textbox(Inches(1.2), Inches(5.5), Inches(11), Inches(0.5))
    p3 = tb3.text_frame.paragraphs[0]
    p3.text = datetime.now().strftime("%Y.%m.%d")
    p3.font.size = Pt(14); p3.font.color.rgb = h("#AAAAAA")

    # === 内容页 ===
    for i, slide in enumerate(slides[:slide_count]):
        sl = prs.slides.add_slide(prs.slide_layouts[6])
        sl.background.fill.solid(); sl.background.fill.fore_color.rgb = h(bg)

        # 顶部色条
        bar = sl.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(0.06))
        bar.fill.solid(); bar.fill.fore_color.rgb = h(sc); bar.line.fill.background()

        # 页码
        nb = sl.shapes.add_shape(1, Inches(0.5), Inches(0.4), Inches(0.55), Inches(0.55))
        nb.fill.solid(); nb.fill.fore_color.rgb = h(sc); nb.line.fill.background()
        np = nb.text_frame.paragraphs[0]; np.text = str(i + 1)
        np.font.size = Pt(14); np.font.bold = True; np.font.color.rgb = h("#FFFFFF")
        np.alignment = PP_ALIGN.CENTER

        # 标题
        tb = sl.shapes.add_textbox(Inches(1.5), Inches(0.3), Inches(11), Inches(0.75))
        tp = tb.text_frame.paragraphs[0]
        tp.text = slide.get("title", f"第{i+1}页"); tp.font.size = Pt(32)
        tp.font.bold = True; tp.font.color.rgb = h(pc); tp.font.name = t["font"]

        # 分隔线
        ln = sl.shapes.add_shape(1, Inches(1.5), Inches(1.2), Inches(2.5), Inches(0.025))
        ln.fill.solid(); ln.fill.fore_color.rgb = h(sc); ln.line.fill.background()

        # 要点
        bullets = slide.get("bullets", ["暂无内容"])
        tb2 = sl.shapes.add_textbox(Inches(1.8), Inches(1.6), Inches(10.5), Inches(5))
        tf2 = tb2.text_frame; tf2.word_wrap = True
        for j, bullet in enumerate(bullets):
            p = tf2.paragraphs[0] if j == 0 else tf2.add_paragraph()
            p.text = f"▸ {bullet}"; p.font.size = Pt(18)
            p.font.color.rgb = h(tc); p.font.name = t["font"]
            p.space_after = Pt(14)

        # 备注
        if slide.get("notes"):
            try:
                sl.notes_slide.notes_text_frame.text = slide["notes"]
            except:
                pass

    output = io.BytesIO()
    prs.save(output); output.seek(0)
    return output


# ==================== 通用去水印 ====================

COMMON_WATERMARKS = [
    "made with gamma", "gamma.app", "gamma", "canva", "wps", "slidesgo",
    "slideshare", "prezi", "beautiful.ai", "slidebean", "powtoon",
    "visme", "piktochart", "genially", "googleslides", "keynote",
    "watermark", "sample", "preview", "draft", "confidential", "demo", "trial",
]


def remove_watermark(file_bytes, filename, target=None):
    temp_dir, extract_dir = tempfile.mkdtemp(), None
    try:
        ip = os.path.join(temp_dir, filename)
        op = os.path.join(temp_dir, f"clean_{filename}")
        with open(ip, 'wb') as f: f.write(file_bytes)
        extract_dir = temp_dir + "_ex"
        with zipfile.ZipFile(ip, 'r') as z: z.extractall(extract_dir)
        keywords = [target.lower()] if target else COMMON_WATERMARKS
        found, removed = set(), 0
        for root_dir, dirs, files in os.walk(extract_dir):
            for file in files:
                if not file.endswith('.xml'): continue
                fp = os.path.join(root_dir, file)
                try:
                    tree = ET.parse(fp); root = tree.getroot(); changed = False
                    for elem in root.iter():
                        if elem.text:
                            for kw in keywords:
                                if kw in elem.text.lower():
                                    found.add(kw); elem.text = ''; changed = True; removed += 1; break
                        for an, av in list(elem.attrib.items()):
                            if av:
                                for kw in keywords:
                                    if kw in str(av).lower():
                                        found.add(kw); del elem.attrib[an]; changed = True; removed += 1; break
                    if changed:
                        tree.write(fp, xml_declaration=True, encoding='UTF-8')
                except ET.ParseError: continue
        with zipfile.ZipFile(op, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root_dir, dirs, files in os.walk(extract_dir):
                for file in files:
                    fp = os.path.join(root_dir, file)
                    zout.write(fp, os.path.relpath(fp, extract_dir))
        with open(op, 'rb') as f: return f.read(), removed, list(found)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        if extract_dir: shutil.rmtree(extract_dir, ignore_errors=True)


# ==================== 布局修复 ====================

def fix_layout(file_bytes, filename):
    temp_dir, extract_dir = tempfile.mkdtemp(), None
    try:
        ip = os.path.join(temp_dir, filename)
        op = os.path.join(temp_dir, f"fixed_{filename}")
        with open(ip, 'wb') as f: f.write(file_bytes)
        extract_dir = temp_dir + "_ex"
        with zipfile.ZipFile(ip, 'r') as z: z.extractall(extract_dir)
        fixed = 0
        for root_dir, dirs, files in os.walk(extract_dir):
            for file in files:
                if not file.endswith('.xml'): continue
                fp = os.path.join(root_dir, file)
                try:
                    tree = ET.parse(fp); root = tree.getroot(); changed = False
                    for elem in root.iter():
                        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
                        if tag in ('rPr', 'defRPr'):
                            for child in elem:
                                ct = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                                if ct == 'latin' and not child.get('typeface', ''):
                                    child.set('typeface', 'Arial'); changed = True
                                elif ct == 'ea' and not child.get('typeface', ''):
                                    child.set('typeface', '微软雅黑'); changed = True
                    if changed: tree.write(fp, xml_declaration=True, encoding='UTF-8'); fixed += 1
                except ET.ParseError: continue
        with zipfile.ZipFile(op, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root_dir, dirs, files in os.walk(extract_dir):
                for file in files:
                    fp = os.path.join(root_dir, file)
                    zout.write(fp, os.path.relpath(fp, extract_dir))
        with open(op, 'rb') as f: return f.read(), fixed
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        if extract_dir: shutil.rmtree(extract_dir, ignore_errors=True)


# ==================== PPT 美化 ====================

def beautify_pptx(file_bytes, filename, theme_name, intensity):
    t = THEMES.get(theme_name, THEMES["专业商务"])
    is_dark = t.get("dark", False)
    pc, sc = t["primary"], t["secondary"]
    bg = "#1E1E1E" if is_dark else "#FFFFFF"
    tc = "#FFFFFF" if is_dark else "#333333"
    font = t["font"]

    def h(c):
        c = c.lstrip("#")
        return RGBColor(int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16))

    prs = Presentation(io.BytesIO(file_bytes))
    changes = 0

    for master in prs.slide_masters:
        try:
            master.background.fill.solid(); master.background.fill.fore_color.rgb = h(bg)
        except: pass

    for slide in prs.slides:
        try:
            slide.background.fill.solid(); slide.background.fill.fore_color.rgb = h(bg)
        except: pass
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        run.font.name = font
                        if intensity != "轻度":
                            if run.font.size and run.font.size >= Pt(24):
                                run.font.color.rgb = h(pc)
                            else:
                                run.font.color.rgb = h(tc)
                        changes += 1

    output = io.BytesIO(); prs.save(output); output.seek(0)
    return output, changes


# ==================== UI ====================

st.title("🪄 PPT 工具箱")
st.caption("AI 智能生成 · 一键美化 · 通用去水印 · 布局修复")

# ---- 侧边栏：权限管理 ----
with st.sidebar:
    st.subheader("🔐 权限管理")

    if not st.session_state.ai_enabled:
        auth_method = st.radio("选择验证方式", ["🔑 管理密码", "🎫 激活码"], horizontal=True)

        if auth_method == "🔑 管理密码":
            pwd = st.text_input("输入管理密码", type="password")
            if st.button("验证密码", use_container_width=True):
                if pwd == ADMIN_PASSWORD:
                    st.session_state.ai_enabled = True
                    st.session_state.auth_type = "admin"
                    st.success("✅ 管理员验证成功！")
                    st.rerun()
                else:
                    st.error("❌ 密码错误")

        else:
            code = st.text_input("输入激活码", placeholder="PPT-XXXX-XXXX")
            if st.button("激活", use_container_width=True):
                entry = validate_code(code)
                if entry:
                    st.session_state.ai_enabled = True
                    st.session_state.auth_type = "activation"
                    st.session_state.code_used = code
                    mark_code_used(code)
                    st.success("✅ 激活成功！")
                    st.rerun()
                else:
                    st.error("❌ 激活码无效或已被使用")
    else:
        if st.session_state.auth_type == "admin":
            st.success("🔓 管理员模式")
        else:
            st.success(f"🔓 已激活: {st.session_state.code_used}")
        st.caption(f"剩余激活码: {sum(1 for c in VALID_CODES if not c['used'])} 个")

    st.divider()
    st.caption(f"API 余额由管理员提供 · DeepSeek")

# ---- 主界面 Tabs ----
tab1, tab2, tab3, tab4 = st.tabs(["🤖 AI 生成 PPT", "✨ 美化 PPT", "🔓 通用去水印", "🔧 布局修复"])

# ======== Tab 1: AI 生成 ========
with tab1:
    st.subheader("一句话或长描述，生成专业 PPT")

    col1, col2, col3 = st.columns(3)
    with col1:
        theme = st.selectbox("🎨 风格", list(THEMES.keys()))
    with col2:
        slide_count = st.slider("📊 页数", 4, 15, 8)
    with col3:
        generation_mode = "AI 智能生成" if st.session_state.ai_enabled else "模板生成（免费）"
        st.info(f"📋 {generation_mode}")

    # 长文本输入区域
    prompt = st.text_area(
        "📝 输入主题或详细描述",
        placeholder="短主题：新能源汽车2025市场分析\n\n也可以粘贴长描述：\n介绍当前新能源汽车市场的整体情况，分析特斯拉、比亚迪、蔚来等主要玩家的竞争格局，重点解读2025年电池技术突破对行业的影响...",
        height=150,
    )

    if st.button("🪄 生成 PPT", type="primary", use_container_width=True):
        if not prompt.strip():
            st.warning("请输入内容")
        else:
            if st.session_state.ai_enabled:
                # AI 模式
                with st.spinner("🤖 AI 正在分析内容并生成 PPT..."):
                    try:
                        raw = call_deepseek(prompt.strip())
                        data = parse_ai_response(raw)
                        slides_count = len(data.get("slides", []))
                        st.success(f"✅ AI 生成成功，共 {slides_count} 页")
                    except Exception as e:
                        st.error(f"AI 生成失败: {e}")
                        st.info("自动切换到模板模式...")
                        outline = OUTLINE_LIBRARY.get(theme, OUTLINE_LIBRARY["专业商务"])
                        data = {
                            "title": prompt.strip()[:60],
                            "subtitle": f"{theme} · 智能生成",
                            "slides": [{"title": item, "bullets": [f"{item}相关内容", "核心要点分析", "实践案例解读"], "notes": ""} for item in outline[:slide_count - 1]],
                        }
            else:
                # 模板模式（免费）
                if not st.session_state.ai_enabled:
                    st.warning("⚠️ AI 生成需要激活码或管理密码，请在左侧边栏解锁。当前使用模板生成。")

                with st.spinner("📋 正在构建 PPT..."):
                    outline = OUTLINE_LIBRARY.get(theme, OUTLINE_LIBRARY["专业商务"])
                    data = {
                        "title": prompt.strip()[:60],
                        "subtitle": f"{theme} · 智能生成",
                        "slides": [{"title": item, "bullets": [f"{item} - 核心内容", f"{item} - 关键分析", f"{item} - 实例说明"], "notes": ""} for item in outline[:slide_count - 1]],
                    }
                    st.success(f"✅ 模板生成成功，共 {len(data['slides']) + 1} 页")

            with st.spinner("📄 构建 PPTX 文件..."):
                pptx_data = build_pptx(data, theme, slide_count)

            # 根据内容生成文件名
            safe_name = re.sub(r'[^\w一-鿿]', '_', prompt.strip()[:30])
            st.download_button("📥 下载 PPTX", pptx_data,
                               file_name=f"{safe_name}_{theme}.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                               use_container_width=True)

# ======== Tab 2: 美化 ========
with tab2:
    st.subheader("✨ PPT 一键美化")
    c1, c2 = st.columns(2)
    with c1:
        b_theme = st.selectbox("🎨 目标风格", list(THEMES.keys()), key="b_theme")
    with c2:
        intensity = st.selectbox("🖌️ 美化程度", ["轻度（字体）", "中度（字体+配色）", "深度（全部重排）"])

    uploaded_b = st.file_uploader("上传 PPTX", type=["pptx"], key="b_upload")
    if uploaded_b:
        if st.button("✨ 开始美化", type="primary", use_container_width=True):
            with st.spinner("美化中..."):
                result, changes = beautify_pptx(uploaded_b.getvalue(), uploaded_b.name, b_theme, intensity)
                st.success(f"✅ 处理了 {changes} 处文本")
                st.download_button("📥 下载美化版", result,
                                   file_name=f"美化_{b_theme}_{uploaded_b.name}",
                                   mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                   use_container_width=True)

# ======== Tab 3: 去水印 ========
with tab3:
    st.subheader("🔓 通用去水印")
    wm = st.text_input("🔍 指定水印文字（可选）", placeholder="留空自动检测 80+ 常见水印")
    uploaded_w = st.file_uploader("上传 PPTX", type=["pptx"], key="w_upload")
    if uploaded_w:
        if st.button("🔓 去水印", use_container_width=True):
            with st.spinner("搜索中..."):
                result, count, found = remove_watermark(uploaded_w.getvalue(), uploaded_w.name,
                                                        wm.strip() if wm.strip() else None)
                if count:
                    st.success(f"✅ 移除 {count} 个标记 ({', '.join(found)})")
                    st.download_button("📥 下载", result, file_name=f"去水印_{uploaded_w.name}",
                                       mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                       use_container_width=True)
                else:
                    st.info("未检测到水印")

# ======== Tab 4: 修复 ========
with tab4:
    st.subheader("🔧 布局修复")
    uploaded_f = st.file_uploader("上传 PPTX", type=["pptx"], key="f_upload")
    if uploaded_f:
        if st.button("🔧 修复", use_container_width=True):
            with st.spinner("修复中..."):
                result, fixed = fix_layout(uploaded_f.getvalue(), uploaded_f.name)
                st.success(f"✅ 修复 {fixed} 处")
                st.download_button("📥 下载修复版", result, file_name=f"修复_{uploaded_f.name}",
                                   mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                   use_container_width=True)
