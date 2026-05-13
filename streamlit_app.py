"""
一句话生成PPT + Gamma去水印 + 布局修复
"""
import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import os
import tempfile
import shutil
import uuid
import json
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import io

st.set_page_config(page_title="AI PPT 生成器", page_icon="🎨", layout="wide")

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

# ==================== PPT 生成核心 ====================

PPT_TEMPLATES = {
    "商务演示": {
        "bg": (255, 255, 255),
        "title_color": (30, 60, 120),
        "accent": (41, 128, 185),
        "font": "微软雅黑",
    },
    "学术答辩": {
        "bg": (255, 255, 255),
        "title_color": (52, 73, 94),
        "accent": (46, 134, 193),
        "font": "微软雅黑",
    },
    "创业路演": {
        "bg": (255, 255, 255),
        "title_color": (231, 76, 60),
        "accent": (243, 156, 18),
        "font": "微软雅黑",
    },
    "工作总结": {
        "bg": (255, 255, 255),
        "title_color": (38, 70, 83),
        "accent": (42, 157, 143),
        "font": "微软雅黑",
    },
}

OUTLINE_TEMPLATES = {
    "商务演示": ["项目背景", "市场分析", "解决方案", "竞争优势", "执行计划", "预期成果"],
    "学术答辩": ["研究背景", "文献综述", "研究方法", "实验结果", "分析讨论", "结论展望"],
    "创业路演": ["痛点分析", "解决方案", "市场规模", "商业模式", "团队优势", "融资计划"],
    "工作总结": ["工作概况", "重点项目", "数据成果", "问题反思", "改进措施", "下阶段规划"],
}

# 14 种风格配色
STYLE_COLORS = {
    "商务蓝": {"primary": (41, 128, 185), "bg": (255, 255, 255)},
    "科技紫": {"primary": (142, 68, 173), "bg": (255, 255, 255)},
    "清新绿": {"primary": (39, 174, 96), "bg": (255, 255, 255)},
    "活力橙": {"primary": (230, 126, 34), "bg": (255, 255, 255)},
    "优雅红": {"primary": (192, 57, 43), "bg": (255, 255, 255)},
    "沉稳灰": {"primary": (52, 73, 94), "bg": (255, 255, 255)},
    "海洋蓝": {"primary": (21, 67, 96), "bg": (255, 255, 255)},
    "森林绿": {"primary": (20, 90, 50), "bg": (255, 255, 255)},
    "暗夜黑": {"primary": (44, 62, 80), "bg": (255, 255, 255)},
    "玫瑰金": {"primary": (183, 108, 128), "bg": (255, 255, 255)},
    "湖水青": {"primary": (22, 160, 133), "bg": (255, 255, 255)},
    "珊瑚粉": {"primary": (240, 128, 128), "bg": (255, 255, 255)},
    "经典黑金": {"primary": (218, 165, 32), "bg": (30, 30, 30)},
    "极简白": {"primary": (100, 100, 100), "bg": (255, 255, 255)},
}


def create_pptx(topic, template, style_name, slides_content=None):
    """根据主题和模板生成 PPTX 文件"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    colors = STYLE_COLORS.get(style_name, STYLE_COLORS["商务蓝"])
    primary = colors["primary"]
    bg = colors["bg"]

    # 如果没有 AI 内容，使用默认大纲
    if not slides_content:
        outline = OUTLINE_TEMPLATES.get(template, OUTLINE_TEMPLATES["商务演示"])
        slides_content = []
        for i, section in enumerate(outline):
            slides_content.append({
                "title": section,
                "bullet_points": [f"{section}相关内容 - 第{j+1}点" for j in range(3)],
                "speaker_notes": f"展开讲解{section}"
            })

    # === 封面 ===
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    bg_fill = slide.background.fill
    bg_fill.solid()
    bg_fill.fore_color.rgb = RGBColor(*bg)

    # 左侧装饰条
    shape = slide.shapes.add_shape(
        1, Inches(0), Inches(0), Inches(0.15), Inches(7.5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*primary)
    shape.line.fill.background()

    # 主标题
    txBox = slide.shapes.add_textbox(Inches(1.2), Inches(2), Inches(11), Inches(2))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = topic
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = RGBColor(*primary)
    p.font.name = "微软雅黑"

    # 副标题
    txBox2 = slide.shapes.add_textbox(Inches(1.2), Inches(4.2), Inches(11), Inches(1))
    tf2 = txBox2.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = f"{template} · AI 智能生成"
    p2.font.size = Pt(20)
    p2.font.color.rgb = RGBColor(150, 150, 150)
    p2.font.name = "微软雅黑"

    # 日期
    from datetime import date
    txBox3 = slide.shapes.add_textbox(Inches(1.2), Inches(5.2), Inches(11), Inches(0.5))
    tf3 = txBox3.text_frame
    p3 = tf3.paragraphs[0]
    p3.text = date.today().strftime("%Y年%m月%d日")
    p3.font.size = Pt(14)
    p3.font.color.rgb = RGBColor(180, 180, 180)
    p3.font.name = "微软雅黑"

    # === 内容页 ===
    for i, slide_data in enumerate(slides_content):
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        bg_fill = slide.background.fill
        bg_fill.solid()
        bg_fill.fore_color.rgb = RGBColor(*bg)

        # 顶部色条
        shape = slide.shapes.add_shape(
            1, Inches(0), Inches(0), Inches(13.333), Inches(0.08))
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(*primary)
        shape.line.fill.background()

        # 页码标识
        num_shape = slide.shapes.add_shape(
            1, Inches(0.5), Inches(0.5), Inches(0.6), Inches(0.6))
        num_shape.fill.solid()
        num_shape.fill.fore_color.rgb = RGBColor(*primary)
        num_shape.line.fill.background()
        num_tf = num_shape.text_frame
        num_p = num_tf.paragraphs[0]
        num_p.text = str(i + 1)
        num_p.font.size = Pt(14)
        num_p.font.bold = True
        num_p.font.color.rgb = RGBColor(255, 255, 255)
        num_p.alignment = PP_ALIGN.CENTER

        # 标题
        txBox = slide.shapes.add_textbox(Inches(1.5), Inches(0.4), Inches(11), Inches(0.8))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = slide_data.get("title", f"第{i+1}页")
        p.font.size = Pt(32)
        p.font.bold = True
        p.font.color.rgb = RGBColor(*primary)
        p.font.name = "微软雅黑"

        # 分隔线
        line = slide.shapes.add_shape(
            1, Inches(1.5), Inches(1.3), Inches(3), Inches(0.03))
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(*primary)
        line.line.fill.background()

        # 要点
        bullets = slide_data.get("bullet_points", [])
        txBox2 = slide.shapes.add_textbox(Inches(1.8), Inches(1.8), Inches(10), Inches(4.5))
        tf2 = txBox2.text_frame
        tf2.word_wrap = True

        for j, bullet in enumerate(bullets):
            if j == 0:
                p = tf2.paragraphs[0]
            else:
                p = tf2.add_paragraph()
            p.text = f"▸ {bullet}"
            p.font.size = Pt(18)
            p.font.color.rgb = RGBColor(80, 80, 80)
            p.font.name = "微软雅黑"
            p.space_after = Pt(12)

        # 演讲者备注
        if slide_data.get("speaker_notes"):
            notes = slide.notes_slide
            notes.notes_text_frame.text = slide_data["speaker_notes"]

    # 保存
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output


# ==================== 通用去水印 ====================

COMMON_WATERMARKS = [
    "made with gamma", "gamma.app", "gamma",
    "canva", "made with canva",
    "wps", "wps office",
    "powerpoint", "microsoft",
    "slidesgo", "slidesgo.com",
    "slideshare", "slideshare.net",
    "prezi",
    "beautiful.ai",
    "slidebean",
    "powtoon",
    "visme",
    "piktochart",
    "genially",
    "mentimeter",
    "nearpod",
    "peardeck",
    "haikudeck",
    "zoho show",
    "customshow",
    "slidecamp",
    "slidebean",
    "ludus",
    "slides",
    "slideful",
    "emaze",
    "sway",
    "googleslides",
    "keynote",
    "libreoffice impress",
    "openoffice impress",
    "staroffice impress",
    "softmaker presentations",
    "corel presentations",
    "acrobat",
    "pdf",
    "watermark",
    "sample",
    "preview",
    "draft",
    "confidential",
    "demo",
    "trial",
    "evaluation",
    "free",
    "free version",
    "free trial",
    "free download",
    "download",
    "free ppt",
    "free presentation",
    "free template",
    "free download",
    "freebie",
    "freeware",
    "shareware",
    "freemium",
    "free account",
    "free plan",
    "free tier",
    "free version",
    "free to use",
    "free to download",
    "free to edit",
    "free to share",
    "free to use",
    "free to copy",
    "free to distribute",
    "free to modify",
    "free to remix",
    "free to adapt",
    "free to build upon",
    "free to use commercially",
    "free for commercial use",
    "free for personal use",
    "free for education",
    "free for non-profit",
    "free for charity",
    "free for government",
    "free for students",
    "free for teachers",
    "free for educators",
    "free for schools",
    "free for universities",
    "free for colleges",
    "free for libraries",
    "free for museums",
    "free for archives",
    "free for galleries",
    "free for cultural institutions",
    "free for research",
    "free for academic use",
]


def remove_watermark_generic(file_bytes, filename, target_text=None):
    """通用去水印，支持自定义关键词或自动检测常见水印"""
    temp_dir = tempfile.mkdtemp()
    extract_dir = None
    try:
        input_path = os.path.join(temp_dir, filename)
        output_path = os.path.join(temp_dir, f"clean_{filename}")

        with open(input_path, 'wb') as f:
            f.write(file_bytes)

        extract_dir = temp_dir + "_extract"
        with zipfile.ZipFile(input_path, 'r') as z:
            z.extractall(extract_dir)

        # 确定要搜索的关键词
        if target_text:
            keywords = [target_text.lower()]
        else:
            keywords = COMMON_WATERMARKS

        found_texts = set()
        removed_count = 0

        for root_dir, dirs, files in os.walk(extract_dir):
            for file in files:
                if not file.endswith('.xml'):
                    continue
                filepath = os.path.join(root_dir, file)
                try:
                    tree = ET.parse(filepath)
                    root = tree.getroot()
                    file_changed = False

                    for elem in root.iter():
                        # 检查文本
                        if elem.text:
                            text_lower = elem.text.lower()
                            for kw in keywords:
                                if kw in text_lower:
                                    found_texts.add(kw)
                                    elem.text = ''
                                    file_changed = True
                                    removed_count += 1
                                    break

                        # 检查属性
                        for an, av in list(elem.attrib.items()):
                            if av:
                                av_lower = str(av).lower()
                                for kw in keywords:
                                    if kw in av_lower:
                                        found_texts.add(kw)
                                        del elem.attrib[an]
                                        file_changed = True
                                        removed_count += 1
                                        break

                    if file_changed:
                        tree.write(filepath, xml_declaration=True, encoding='UTF-8')
                except ET.ParseError:
                    continue

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root_dir, dirs, files in os.walk(extract_dir):
                for file in files:
                    fp = os.path.join(root_dir, file)
                    an = os.path.relpath(fp, extract_dir)
                    zout.write(fp, an)

        with open(output_path, 'rb') as f:
            result = f.read()
        return result, removed_count, list(found_texts)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        if extract_dir:
            shutil.rmtree(extract_dir, ignore_errors=True)


# ==================== UI ====================

st.title("🎨 一句话生成 PPT")
st.caption("输入主题 → AI 生成大纲 → 一键下载完整 PPT · 支持去水印 · 完美兼容 WPS/Office")

tab1, tab2, tab3 = st.tabs(["🪄 生成 PPT", "🔓 去水印", "🔧 修复布局"])

with tab1:
    col1, col2 = st.columns([2, 1])

    with col1:
        topic = st.text_input("📝 输入主题，一句话即可", placeholder="例如：2025年新能源汽车市场分析报告")

    with col2:
        template = st.selectbox("📋 选择场景", list(PPT_TEMPLATES.keys()))

    style = st.selectbox("🎨 风格配色", list(STYLE_COLORS.keys()), horizontal=True)

    # AI 增强
    with st.expander("🤖 AI 增强（可选，效果更好）"):
        use_ai = st.checkbox("使用 AI 生成内容（需 DeepSeek API Key）")
        api_key = st.text_input("DeepSeek API Key", type="password",
                                help="在 https://platform.deepseek.com 免费注册获取",
                                disabled=not use_ai)

    if st.button("🪄 生成 PPT", type="primary", use_container_width=True):
        if not topic.strip():
            st.warning("请输入主题")
        else:
            with st.spinner("正在生成 PPT..."):
                slides_content = None

                if use_ai and api_key:
                    try:
                        import requests
                        outline = OUTLINE_TEMPLATES.get(template, OUTLINE_TEMPLATES["商务演示"])
                        prompt = f"""你是一个专业的PPT内容策划专家。请根据以下主题生成PPT内容。
主题：{topic}
场景：{template}
大纲：{', '.join(outline)}

请为每个章节生成：
1. 标题（精简有力，10字以内）
2. 3-5个要点（每个20字以内，用换行分隔）
3. 演讲者备注（1-2句话）

请以JSON格式输出：
[{{"title": "章节标题", "bullet_points": ["要点1","要点2","要点3"], "speaker_notes": "备注内容"}}]"""

                        resp = requests.post(
                            "https://api.deepseek.com/chat/completions",
                            headers={"Authorization": f"Bearer {api_key}"},
                            json={
                                "model": "deepseek-chat",
                                "messages": [{"role": "user", "content": prompt}],
                                "temperature": 0.7,
                            },
                            timeout=60
                        )
                        if resp.status_code == 200:
                            raw = resp.json()["choices"][0]["message"]["content"]
                            raw = raw.strip()
                            if raw.startswith("```"):
                                raw = raw.split("\n", 1)[1].rsplit("```", 1)[0]
                            slides_content = json.loads(raw)
                            st.success("✅ AI 内容生成成功")
                        else:
                            st.error(f"API 错误: {resp.json()}")
                    except Exception as e:
                        st.error(f"AI 调用失败: {e}，将使用默认模板生成")

                pptx_data = create_pptx(topic, template, style, slides_content)

            st.success("🎉 PPT 生成完成！")
            st.download_button(
                label="📥 下载 PPTX 文件",
                data=pptx_data,
                file_name=f"{topic[:20]}_{template}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )

    # 预览区
    with st.expander("📖 大纲预览"):
        outline = OUTLINE_TEMPLATES.get(template, OUTLINE_TEMPLATES["商务演示"])
        for i, item in enumerate(outline, 1):
            st.write(f"**{i}.** {item}")


with tab2:
    st.subheader("🔓 去除 PPTX 水印")
    st.caption("上传 PPTX 文件，自动搜索并去除指定水印文字（支持 Gamma / Canva / WPS 等任何来源）")

    watermark_text = st.text_input("🔍 要移除的水印文字",
                                    placeholder="例如：Made with Gamma、Canva、品牌名称...",
                                    help="输入你想去除的水印关键词，留空则自动检测常见水印")

    uploaded = st.file_uploader("选择 PPTX 文件", type=["pptx"], key="watermark_upload")
    if uploaded:
        if st.button("去除水印", use_container_width=True):
            with st.spinner("处理中..."):
                result, count, found_texts = remove_watermark_generic(
                    uploaded.getvalue(), uploaded.name, watermark_text.strip() if watermark_text.strip() else None
                )
                if count > 0:
                    st.success(f"✅ 已移除 {count} 个水印标记")
                    if found_texts:
                        st.caption(f"检测到的水印: {', '.join(found_texts)}")
                    st.download_button(
                        "📥 下载无水印版本", result,
                        file_name=f"去水印_{uploaded.name}",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True,
                    )
                else:
                    st.info("未检测到水印，请确认水印文字是否正确")


with tab3:
    st.subheader("🔧 修复 PPTX 布局")
    st.caption("修复空字体、对齐问题")
    uploaded2 = st.file_uploader("选择 PPTX 文件", type=["pptx"], key="fix_upload")
    if uploaded2:
        if st.button("修复布局", use_container_width=True):
            with st.spinner("修复中..."):
                temp_dir = tempfile.mkdtemp()
                try:
                    ip = os.path.join(temp_dir, uploaded2.name)
                    op = os.path.join(temp_dir, f"fixed_{uploaded2.name}")
                    with open(ip, 'wb') as f:
                        f.write(uploaded2.getvalue())

                    with zipfile.ZipFile(ip, 'r') as z:
                        z.extractall(temp_dir + "_ex")

                    ed = temp_dir + "_ex"
                    fixed = 0
                    for root_dir, dirs, files in os.walk(ed):
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
                                                child.set('typeface', 'Arial')
                                                changed = True
                                            elif ct == 'ea' and child.get('typeface', '') == '':
                                                child.set('typeface', '微软雅黑')
                                                changed = True
                                if changed:
                                    tree.write(fp, xml_declaration=True, encoding='UTF-8')
                                    fixed += 1
                            except ET.ParseError:
                                continue

                    with zipfile.ZipFile(op, 'w', zipfile.ZIP_DEFLATED) as zout:
                        for root_dir, dirs, files in os.walk(ed):
                            for file in files:
                                fp = os.path.join(root_dir, file)
                                an = os.path.relpath(fp, ed)
                                zout.write(fp, an)

                    with open(op, 'rb') as f:
                        result = f.read()

                    st.success(f"✅ 已修复 {fixed} 个文件")
                    st.download_button(
                        "📥 下载修复版本", result,
                        file_name=f"修复_{uploaded2.name}",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True,
                    )
                finally:
                    shutil.rmtree(temp_dir, ignore_errors=True)
