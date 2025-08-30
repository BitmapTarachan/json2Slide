from typing import List
from pptx.util import Pt

def render_content_default(factory, data):
    s = factory._new_slide(data)
    factory._add_slide_title(s, data.get("title",""))

    # Subhead
    subhead = data.get("subhead")
    if subhead:
        s_rect = factory.layout.get_rect("contentSlide.subhead")
        sbox = s.shapes.add_textbox(s_rect["left"], s_rect["top"], s_rect["width"], s_rect["height"])
        sp = sbox.text_frame.paragraphs[0]
        factory._style_text(
            sp,
            subhead,
            factory.fonts["sizes"]["subhead"],
            color=factory.colors["subtext"]
        )
        sbox.text_frame.word_wrap = True

    # Body area
    b_rect = factory.layout.get_rect("contentSlide.body")

    points: List[str] = data.get("points", [])
    body_text: str = data.get("bodyText", "")

    last_y = b_rect["top"]

    # --- 箇条書き ---
    if points:
        # 箇条書き部分
        line_spacing = factory.fonts["sizes"]["body"].pt + 10  # フォントサイズ + 行間
        bbox = s.shapes.add_textbox(b_rect["left"], b_rect["top"], b_rect["width"], b_rect["height"])
        tf = bbox.text_frame
        tf.clear()
        tf.word_wrap = True

        for i, line in enumerate(points):
            p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
            factory._style_text(p, line, factory.fonts["sizes"]["body"], color=factory.colors["text"])
            p.space_after = Pt(10)

        # 箇条書きの下位置を行数で見積もり（cm換算）
        from pptx.util import Cm
        lines_used = len(points)
        last_y = b_rect["top"] + Cm((line_spacing * lines_used) / 28.35)

    # --- 長文 ---
    if body_text:
        lbox = s.shapes.add_textbox(b_rect["left"], last_y + Pt(20), b_rect["width"], b_rect["height"])
        tf2 = lbox.text_frame
        tf2.word_wrap = True  # 折り返し有効
        lp = tf2.paragraphs[0]
        factory._style_text(lp, body_text, factory.fonts["sizes"]["body"], color=factory.colors["text"])

    return s
