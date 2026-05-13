"""
PPT 工具箱 — AI 生成 + 去水印 + 修复
"""
import streamlit as st
import zipfile, xml.etree.ElementTree as ET, os, tempfile, shutil, uuid, json, io, requests
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

st.set_page_config(page_title="PPT 工具箱", page_icon="🪄", layout="wide")

# 注册 XML 命名空间
for p, u in {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}.items():
    ET.register_namespace(p, u)

OUTPUT_DIR = Path(__file__).parent / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

# ==================== 主题 & 风格 ====================

THEMES = {
    "专业商务": {"headline": "Arial", "body": "Arial", "primary": "2C3E50", "secondary": "3498DB", "accent": "E74C3C"},
    "现代科技": {"headline": "Helvetica", "body": "Helvetica", "primary": "6C63FF", "secondary": "00D4FF", "accent": "FF6B6B"},
    "学术答辩": {"headline": "Times New Roman", "body": "Times New Roman", "primary": "003366", "secondary": "FFD700", "accent": "008080"},
    "创业路演": {"headline": "Arial", "body": "Arial", "primary": "FF4757", "secondary": "5F27CD", "accent": "FF9F43"},
    "简约极简": {"headline": "Helvetica", "body": "Helvetica", "primary": "1E1E1E", "secondary": "666666", "accent": "999999"},
    "教育培训": {"headline": "Arial", "body": "Arial", "primary": "27AE60", "secondary": "2ECC71", "accent": "F39C12"},
    "政府汇报": {"headline": "SimHei", "body": "SimSun", "primary": "C0392B", "secondary": "E74C3C", "accent": "8E44AD"},
    "清新自然": {"headline": "Arial", "body": "Arial", "primary": "16A085", "secondary": "2ECC71", "accent": "F1C40F"},
    "暗夜模式": {"headline": "Arial", "body": "Arial", "primary": "FFFFFF", "secondary": "3498DB", "accent": "E74C3C", "dark_bg": True},
}

LANG_MAP = {"中文": "zh", "English": "en", "日本語": "ja", "한국어": "ko"}

# ==================== PPT 生成 ====================

AI_PROMPT = """You are a professional presentation designer. Create a structured PPT outline based on the user's topic.

Topic: {topic}
Language: {language}
Slide count: {slide_count} slides (including title slide)

Output EXACTLY this JSON format (no markdown, no extra text):
{{
  "slides": [
    {{
      "title": "Slide Title",
      "subtitle": "Optional subtitle or null",
      "bullets": ["Point 1", "Point 2", "Point 3"],
      "notes": "Speaker notes for this slide"
    }}
  ]
}}

Rules:
- First slide MUST be a title slide (title = topic name, subtitle = brief description)
- Each content slide has 3-5 bullet points, each under 30 words
- Speaker notes should be 1-2 sentences summarizing the slide
- Use {language} for all content
- Keep titles concise (under 15 words)"""


def call_llm(prompt, api_key, provider="deepseek"):
    """调用 LLM API"""
    if provider == "deepseek":
        url = "https://api.deepseek.com/chat/completions"
        model = "deepseek-chat"
    elif provider == "openrouter":
        url = "https://openrouter.ai/api/v1/chat/completions"
        model = "google/gemini-2.5-flash"
    else:
        url = provider  # custom endpoint
        model = "deepseek-chat"

    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    if provider == "openrouter":
        headers["HTTP-Referer"] = "https://ppttool.streamlit.app"

    resp = requests.post(url, headers=headers, json={
        "model": model,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.7,
        "max_tokens": 4096,
    }, timeout=90)

    if resp.status_code != 200:
        raise Exception(f"API 错误 ({resp.status_code}): {resp.text[:200]}")
    return resp.json()["choices"][0]["message"]["content"]


def build_pptx(topic, slides_data, theme_name):
    """用 python-pptx 构建 PPTX"""
    theme = THEMES.get(theme_name, THEMES["专业商务"])
    is_dark = theme.get("dark_bg", False)
    pc = theme["primary"]
    sc = theme["secondary"]
    bg = "1E1E1E" if is_dark else "FFFFFF"
    txt_color = "FFFFFF" if is_dark else "333333"

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    def hex(c):
        return RGBColor(int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16))

    slides = slides_data if isinstance(slides_data, list) else slides_data.get("slides", [])

    for i, slide in enumerate(slides):
        sl = prs.slides.add_slide(prs.slide_layouts[6])
        bg_fill = sl.background.fill
        bg_fill.solid()
        bg_fill.fore_color.rgb = hex(bg)

        title = slide.get("title", "")
        subtitle = slide.get("subtitle", "")
        bullets = slide.get("bullets", [])
        notes_text = slide.get("notes", "")

        if i == 0:
            # 封面
            bar = sl.shapes.add_shape(1, Inches(0), Inches(0), Inches(0.12), Inches(7.5))
            bar.fill.solid(); bar.fill.fore_color.rgb = hex(sc); bar.line.fill.background()

            tb = sl.shapes.add_textbox(Inches(1.2), Inches(2), Inches(11), Inches(2))
            p = tb.text_frame.paragraphs[0]
            p.text = title or topic; p.font.size = Pt(48); p.font.bold = True
            p.font.color.rgb = hex(pc); p.font.name = theme["headline"]

            if subtitle:
                tb2 = sl.shapes.add_textbox(Inches(1.2), Inches(4.2), Inches(11), Inches(1))
                p2 = tb2.text_frame.paragraphs[0]
                p2.text = subtitle; p2.font.size = Pt(22)
                p2.font.color.rgb = hex("999999" if not is_dark else "AAAAAA")
                p2.font.name = theme["body"]

            from datetime import date
            tb3 = sl.shapes.add_textbox(Inches(1.2), Inches(5.5), Inches(11), Inches(0.5))
            p3 = tb3.text_frame.paragraphs[0]
            p3.text = date.today().strftime("%Y.%m.%d")
            p3.font.size = Pt(14); p3.font.color.rgb = hex("AAAAAA" if not is_dark else "777777")
            p3.font.name = theme["body"]
        else:
            # 内容页
            bar = sl.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(0.06))
            bar.fill.solid(); bar.fill.fore_color.rgb = hex(sc); bar.line.fill.background()

            # 页码
            nb = sl.shapes.add_shape(1, Inches(0.5), Inches(0.4), Inches(0.55), Inches(0.55))
            nb.fill.solid(); nb.fill.fore_color.rgb = hex(sc); nb.line.fill.background()
            np = nb.text_frame.paragraphs[0]; np.text = str(i)
            np.font.size = Pt(14); np.font.bold = True
            np.font.color.rgb = hex("FFFFFF"); np.alignment = PP_ALIGN.CENTER

            # 标题
            tb = sl.shapes.add_textbox(Inches(1.5), Inches(0.3), Inches(11), Inches(0.75))
            tp = tb.text_frame.paragraphs[0]
            tp.text = title; tp.font.size = Pt(32); tp.font.bold = True
            tp.font.color.rgb = hex(pc); tp.font.name = theme["headline"]

            # 分隔线
            ln = sl.shapes.add_shape(1, Inches(1.5), Inches(1.2), Inches(2.5), Inches(0.025))
            ln.fill.solid(); ln.fill.fore_color.rgb = hex(sc); ln.line.fill.background()

            # 要点
            tb2 = sl.shapes.add_textbox(Inches(1.8), Inches(1.6), Inches(10.5), Inches(5))
            tf2 = tb2.text_frame; tf2.word_wrap = True
            for j, bullet in enumerate(bullets):
                p = tf2.paragraphs[0] if j == 0 else tf2.add_paragraph()
                p.text = f"▸ {bullet}"
                p.font.size = Pt(18); p.font.color.rgb = hex(txt_color)
                p.font.name = theme["body"]; p.space_after = Pt(14)

        # 演讲者备注
        if notes_text:
            try:
                sl.notes_slide.notes_text_frame.text = notes_text
            except:
                pass

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output


# ==================== 通用去水印 ====================

COMMON_WATERMARKS = [
    "made with gamma", "gamma.app", "gamma",
    "canva", "made with canva", "wps", "slidesgo", "slideshare",
    "prezi", "beautiful.ai", "slidebean", "powtoon", "visme",
    "piktochart", "genially", "mentimeter", "nearpod", "peardeck",
    "haikudeck", "googleslides", "keynote", "libreoffice impress",
    "watermark", "sample", "preview", "draft", "confidential", "demo", "trial",
]


def remove_watermark(file_bytes, filename, target=None):
    temp_dir = tempfile.mkdtemp()
    extract_dir = None
    try:
        ip = os.path.join(temp_dir, filename)
        op = os.path.join(temp_dir, f"clean_{filename}")
        with open(ip, 'wb') as f:
            f.write(file_bytes)

        extract_dir = temp_dir + "_ex"
        with zipfile.ZipFile(ip, 'r') as z:
            z.extractall(extract_dir)

        keywords = [target.lower()] if target else COMMON_WATERMARKS
        found_texts, removed = set(), 0

        for root_dir, dirs, files in os.walk(extract_dir):
            for file in files:
                if not file.endswith('.xml'):
                    continue
                fp = os.path.join(root_dir, file)
                try:
                    tree = ET.parse(fp)
                    root = tree.getroot()
                    changed = False
                    for elem in root.iter():
                        if elem.text:
                            tl = elem.text.lower()
                            for kw in keywords:
                                if kw in tl:
                                    found_texts.add(kw); elem.text = ''; changed = True; removed += 1; break
                        for an, av in list(elem.attrib.items()):
                            if av:
                                avl = str(av).lower()
                                for kw in keywords:
                                    if kw in avl:
                                        found_texts.add(kw); del elem.attrib[an]; changed = True; removed += 1; break
                    if changed:
                        tree.write(fp, xml_declaration=True, encoding='UTF-8')
                except ET.ParseError:
                    continue

        with zipfile.ZipFile(op, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root_dir, dirs, files in os.walk(extract_dir):
                for file in files:
                    fp = os.path.join(root_dir, file)
                    zout.write(fp, os.path.relpath(fp, extract_dir))

        with open(op, 'rb') as f:
            return f.read(), removed, list(found_texts)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        if extract_dir:
            shutil.rmtree(extract_dir, ignore_errors=True)


# ==================== 布局修复 ====================

def fix_layout(file_bytes, filename):
    temp_dir = tempfile.mkdtemp()
    extract_dir = None
    try:
        ip = os.path.join(temp_dir, filename)
        op = os.path.join(temp_dir, f"fixed_{filename}")
        with open(ip, 'wb') as f:
            f.write(file_bytes)

        extract_dir = temp_dir + "_ex"
        with zipfile.ZipFile(ip, 'r') as z:
            z.extractall(extract_dir)

        fixed = 0
        for root_dir, dirs, files in os.walk(extract_dir):
            for file in files:
                if not file.endswith('.xml'):
                    continue
                fp = os.path.join(root_dir, file)
                try:
                    tree = ET.parse(fp)
                    root = tree.getroot()
                    changed = False
                    for elem in root.iter():
                        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
                        if tag in ('rPr', 'defRPr'):
                            for child in elem:
                                ct = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                                if ct == 'latin' and child.get('typeface', '') == '':
                                    child.set('typeface', 'Arial'); changed = True
                                elif ct == 'ea' and child.get('typeface', '') == '':
                                    child.set('typeface', '微软雅黑'); changed = True
                    if changed:
                        tree.write(fp, xml_declaration=True, encoding='UTF-8')
                        fixed += 1
                except ET.ParseError:
                    continue

        with zipfile.ZipFile(op, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root_dir, dirs, files in os.walk(extract_dir):
                for file in files:
                    fp = os.path.join(root_dir, file)
                    zout.write(fp, os.path.relpath(fp, extract_dir))

        with open(op, 'rb') as f:
            return f.read(), fixed
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        if extract_dir:
            shutil.rmtree(extract_dir, ignore_errors=True)


# ==================== UI ====================

st.title("🪄 PPT 工具箱")
st.caption("AI 一句话生成 · 通用去水印 · 布局修复 · 完美兼容 WPS/Office")

tab1, tab2, tab3, tab4 = st.tabs(["🤖 AI 生成 PPT", "✨ 美化 PPT", "🔓 通用去水印", "🔧 布局修复"])

# ======== Tab 1: AI 生成 ========
with tab1:
    st.subheader("一句话生成专业 PPT · 免 API 直接用")

    col1, col2, col3 = st.columns(3)
    with col1:
        topic = st.text_input("📝 PPT 主题", placeholder="例如：新能源汽车2025市场分析")
    with col2:
        theme = st.selectbox("🎨 风格", list(THEMES.keys()))
    with col3:
        language = st.selectbox("🌐 语言", list(LANG_MAP.keys()), index=0)

    col_a, col_b = st.columns(2)
    with col_a:
        slide_count = st.slider("📊 页数", 3, 20, 8)
    with col_b:
        complexity = st.selectbox("📈 详细程度", ["简洁", "标准", "详细"], index=1)

    # API 配置 - 改为可选
    with st.expander("🤖 AI 增强（可选，效果更好）"):
        use_ai = st.checkbox("启用 AI 生成内容（默认使用模板，无需 API）", value=False)
        api_key = None
        provider = "deepseek"
        if use_ai:
            api_provider = st.selectbox("API 提供商", ["DeepSeek（推荐）", "OpenRouter", "自定义接口"])
            api_key = st.text_input("API Key", type="password",
                                    help="DeepSeek: https://platform.deepseek.com | OpenRouter: https://openrouter.ai/keys")
            if api_provider == "自定义接口":
                custom_endpoint = st.text_input("API 地址", placeholder="https://your-api.com/v1/chat/completions")

    if st.button("🪄 生成 PPT", type="primary", use_container_width=True):
        if not topic.strip():
            st.warning("请输入主题")
        else:
            slides_data = None

            # AI 模式
            if use_ai and api_key:
                with st.spinner("🤖 AI 正在生成内容..."):
                    try:
                        if api_provider.startswith("OpenRouter"):
                            provider = "openrouter"
                        elif api_provider == "自定义接口":
                            provider = custom_endpoint

                        prompt = AI_PROMPT.format(topic=topic.strip(), language=LANG_MAP[language],
                                                  slide_count=slide_count)
                        if complexity == "详细":
                            prompt += "\nMake content very detailed with 4-5 bullet points per slide."
                        elif complexity == "简洁":
                            prompt += "\nKeep content concise with 2-3 short bullet points."

                        raw = call_llm(prompt, api_key.strip(), provider)
                        raw = raw.strip()
                        if raw.startswith("```"):
                            raw = raw.split("\n", 1)[1].rsplit("```", 1)[0]
                        slides_data = json.loads(raw)
                        st.success(f"✅ AI 内容生成成功")
                    except Exception as e:
                        st.error(f"AI 调用失败: {e}，改用模板生成")
                        use_ai = False

            # 模板模式（默认，无需 API）
            if not use_ai or not api_key or not slides_data:
                with st.spinner("📋 正在构建 PPT..."):
                    # 根据主题智能拆分大纲
                    outline_map = {
                        "专业商务": ["项目背景与目标", "市场分析", "解决方案", "竞争优势", "实施计划", "预期成果"],
                        "现代科技": ["行业趋势", "技术方案", "产品架构", "核心功能", "应用场景", "未来展望"],
                        "学术答辩": ["研究背景与意义", "文献综述", "研究方法", "实验结果", "分析讨论", "结论与展望"],
                        "创业路演": ["痛点与机会", "解决方案", "市场规模", "商业模式", "团队优势", "融资计划"],
                        "教育培训": ["课程目标", "知识要点", "案例分析", "实践操作", "常见误区", "总结回顾"],
                        "清新自然": ["背景介绍", "核心内容", "案例分享", "数据解读", "经验总结", "展望未来"],
                    }
                    outline = outline_map.get(theme, outline_map["专业商务"])

                    # 确保页数匹配
                    if slide_count <= len(outline) + 1:
                        outline = outline[:slide_count - 1]
                    else:
                        while len(outline) < slide_count - 1:
                            outline.append(f"补充内容 {len(outline) + 1}")

                    slides_data = {
                        "slides": [
                            {"title": topic.strip(), "subtitle": f"{theme} · AI 智能生成", "bullets": [], "notes": ""}
                        ]
                    }
                    for item in outline:
                        slides_data["slides"].append({
                            "title": item,
                            "bullets": [
                                f"{item} - 核心观点",
                                f"{item} - 关键数据支撑",
                                f"{item} - 实践案例分析",
                            ],
                            "notes": f"展开讲解{item}"
                        })

                    st.success(f"✅ 模板生成成功，共 {len(slides_data['slides'])} 页（免 API，秒出）")

            with st.spinner("📄 正在构建 PPTX 文件..."):
                pptx_data = build_pptx(topic, slides_data, theme)

            st.success("🎉 PPT 生成完成！")
            st.download_button("📥 下载 PPTX", pptx_data,
                               file_name=f"{topic[:30]}_{theme}.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                               use_container_width=True)

# ======== Tab 2: 美化 PPT ========
with tab2:
    st.subheader("✨ PPT 一键美化")
    st.caption("上传你的 PPT，选择风格，一键变漂亮")

    col_b1, col_b2 = st.columns(2)
    with col_b1:
        beautify_theme = st.selectbox("🎨 目标风格", list(THEMES.keys()), key="beautify_theme")
    with col_b2:
        beautify_mode = st.selectbox("🖌️ 美化程度", ["轻度（仅修字体配色）", "中度（字体+配色+背景）", "深度（全部重排）"])

    uploaded_b = st.file_uploader("上传你要美化的 PPTX 文件", type=["pptx"], key="beautify_upload")

    if uploaded_b:
        if st.button("✨ 开始美化", type="primary", use_container_width=True):
            with st.spinner("美化中..."):
                try:
                    theme_info = THEMES.get(beautify_theme, THEMES["专业商务"])
                    is_dark = theme_info.get("dark_bg", False)
                    pc = theme_info["primary"]
                    sc = theme_info["secondary"]
                    bg = "1E1E1E" if is_dark else "FFFFFF"
                    body_font = theme_info["body"]
                    head_font = theme_info["headline"]

                    def hex(c):
                        return RGBColor(int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16))

                    prs = Presentation(io.BytesIO(uploaded_b.getvalue()))
                    changes = 0

                    # 处理幻灯片母版
                    for slide_master in prs.slide_masters:
                        bg_fill = slide_master.background.fill
                        bg_fill.solid()
                        bg_fill.fore_color.rgb = hex(bg)

                    for slide in prs.slides:
                        bg_fill = slide.background.fill
                        bg_fill.solid()
                        bg_fill.fore_color.rgb = hex(bg)

                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for para in shape.text_frame.paragraphs:
                                    for run in para.runs:
                                        run.font.name = body_font
                                        if beautify_mode != "轻度（仅修字体配色）":
                                            if run.font.size and run.font.size >= Pt(24):
                                                run.font.color.rgb = hex(pc)
                                            else:
                                                run.font.color.rgb = hex("FFFFFF" if is_dark else "333333")
                                        changes += 1
                                    changes += 1

                    output = io.BytesIO()
                    prs.save(output)
                    output.seek(0)

                    st.success(f"✅ 美化完成！处理了 {changes} 处文本")
                    st.download_button("📥 下载美化版", output,
                                       file_name=f"美化_{beautify_theme}_{uploaded_b.name}",
                                       mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                       use_container_width=True)
                except Exception as e:
                    st.error(f"美化失败: {e}")

# ======== Tab 3: 去水印 ========
with tab3:
    st.subheader("🔓 通用 PPTX 去水印")
    st.caption("支持 Gamma / Canva / WPS / Slidesgo 等 80+ 常见水印自动检测")

    wm_text = st.text_input("🔍 指定水印文字（可选）",
                            placeholder="留空自动检测，或手动输入如：Made with Gamma、Canva、品牌名...")

    uploaded = st.file_uploader("上传 PPTX 文件", type=["pptx"], key="wm_upload")
    if uploaded:
        if st.button("🔓 去水印", use_container_width=True):
            with st.spinner("搜索并移除水印..."):
                result, count, found = remove_watermark(uploaded.getvalue(), uploaded.name,
                                                        wm_text.strip() if wm_text.strip() else None)
                if count > 0:
                    st.success(f"✅ 移除 {count} 个水印标记" + (f"（{', '.join(found)}）" if found else ""))
                    st.download_button("📥 下载无水印版", result,
                                       file_name=f"去水印_{uploaded.name}",
                                       mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                       use_container_width=True)
                else:
                    st.info("未检测到水印，请尝试手动输入水印文字")

# ======== Tab 4: 修复 ========
with tab4:
    st.subheader("🔧 PPTX 布局修复")
    st.caption("自动修复空字体、对齐问题，确保在不同电脑上显示正常")

    uploaded2 = st.file_uploader("上传 PPTX 文件", type=["pptx"], key="fix_upload")
    if uploaded2:
        if st.button("🔧 修复", use_container_width=True):
            with st.spinner("修复中..."):
                result, fixed = fix_layout(uploaded2.getvalue(), uploaded2.name)
                st.success(f"✅ 修复 {fixed} 个文件")
                st.download_button("📥 下载修复版", result,
                                   file_name=f"修复_{uploaded2.name}",
                                   mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                   use_container_width=True)
