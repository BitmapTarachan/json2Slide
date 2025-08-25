# -*- coding: utf-8 -*-
"""
majin-style slide factory with python-pptx
- AIは「設計図(JSON)」のみ生成
- この工場がテンプレートに流し込み、PPTXを安定生成
- 主要スライド型: title / section / content / two_column / compare / quote /
                  process / timeline / image / table / key_takeaways
"""
import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Cm


# -------- ユーザー環境に合わせて調整可能な既定値 --------
DEFAULT_FONT = "Biz UDゴシック"           # 日本語フォントを既定化
TITLE_FONT_SIZE = Pt(36)
SUBTITLE_FONT_SIZE = Pt(18)
HEADING_FONT_SIZE = Pt(28)
BODY_FONT_SIZE = Pt(18)
CAPTION_FONT_SIZE = Pt(14)

# レイアウトの既定マップ（テンプレートにより差異あり）
# 必要に応じて調整してください。
DEFAULT_LAYOUT_MAP = {
    "title": 0,           # Title Slide
    "content": 1,         # Title and Content
    "section": 2,         # Section Header
    "two_content": 3,     # Two Content
    "comparison": 4,      # Comparison
    "title_only": 5,      # Title Only
    "blank": 6,           # Blank
    "with_caption": 7,    # Content with Caption
    "pic_with_caption": 8 # Picture with Caption
}

# ---------------- Google風 CONFIG ----------------
CONFIG = {
    "BASE_PX": {"W": 960, "H": 540},

    # ---------------- テーマカラー定義 ----------------
    "COLORS": {
        "primary": RGBColor(0x42, 0x85, 0xF4),      # メインカラー（タイトル・強調）
        "accent": RGBColor(0xFB, 0xBC, 0x04),       # アクセントカラー（ハイライト・バー）
        "background": RGBColor(0xFF, 0xFF, 0xFF),   # スライド背景
        "surface": RGBColor(0xF8, 0xF9, 0xFA),      # セクション背景・ボックス背景
        "text": RGBColor(0x33, 0x33, 0x33),         # 標準本文
        "subtext": RGBColor(0x9E, 0x9E, 0x9E),      # 補助テキスト（日付・キャプション）
        "ghost": RGBColor(0xEF, 0xEF, 0xED),        # ゴースト数字・区切り用
    },

    # ---------------- フォント定義 ----------------
    "FONTS": {
        "family": "Biz UDゴシック",
        "sizes": {
            "title": Pt(45),
            "sectionTitle": Pt(38),
            "contentTitle": Pt(28),
            "subhead": Pt(18),
            "body": Pt(16),
            "caption": Pt(12),
            "ghostNum": Pt(180),
        }
    },

    # ---------------- レイアウト座標 ----------------
    "POS_PX": {
        "titleSlide": {
            "title": {"left": 50, "top": 230, "width": 800, "height": 90},
            "date": {"left": 50, "top": 340, "width": 250, "height": 40},
        },
        "sectionSlide": {
            "title": {"left": 55, "top": 230, "width": 840, "height": 80},
            "ghostNum": {"left": 35, "top": 120, "width": 300, "height": 200},
        },
        "contentSlide": {
            "title": {"left": 25, "top": 60, "width": 830, "height": 65},
            "subhead": {"left": 25, "top": 140, "width": 830, "height": 30},
            "body": {"left": 25, "top": 172, "width": 910, "height": 303},
        }
    }
}

# ---------------- Layout Manager ----------------
class LayoutManager:
    def __init__(self, config):
        self.cfg = config
        self.base_w = config["BASE_PX"]["W"] * 0.75
        self.base_h = config["BASE_PX"]["H"] * 0.75
        self.page_w = Inches(13.33)  # 16:9 幅
        self.page_h = Inches(7.5)    # 16:9 高さ
        self.scale_x = self.page_w / self.base_w
        self.scale_y = self.page_h / self.base_h

    def get_rect(self, path: str):
        keys = path.split(".")
        pos = self.cfg["POS_PX"]
        for k in keys:
            pos = pos[k]
        def px2pt(px): return px * 0.75
        return {
            "left": px2pt(pos["left"]) * self.scale_x,
            "top": px2pt(pos["top"]) * self.scale_y,
            "width": px2pt(pos["width"]) * self.scale_x,
            "height": px2pt(pos["height"]) * self.scale_y,
        }

# ---------------------- ユーティリティ ----------------------
def set_paragraph_style(paragraph, text: str, font_size: Pt, bold=False, italic=False, color: Optional[RGBColor]=None, align=None):
    paragraph.text = text
    run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
    run.font.name = DEFAULT_FONT
    run.font.size = font_size
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color
    if align:
        paragraph.alignment = align


def set_text_frame_bullets(tf, lines: List[str], level: int = 0):
    """最初の段落を上書きし、以降は追加。"""
    tf.clear()
    if not lines:
        return
    p0 = tf.paragraphs[0]
    set_paragraph_style(p0, lines[0], BODY_FONT_SIZE)
    p0.level = level
    for line in lines[1:]:
        p = tf.add_paragraph()
        set_paragraph_style(p, line, BODY_FONT_SIZE)
        p.level = level


def add_speaker_notes(slide, notes: Optional[str]):
    if not notes:
        return
    ns = slide.notes_slide
    tf = ns.notes_text_frame
    tf.clear()
    p = tf.paragraphs[0]
    set_paragraph_style(p, notes, Pt(14))


def ensure_list(x) -> List[Any]:
    if x is None:
        return []
    return x if isinstance(x, list) else [x]


# ---------------- Slide Factory ----------------
class SlideFactory:
    def __init__(self, config=CONFIG):
        self.config = config
        self.colors = config["COLORS"]
        self.fonts = config["FONTS"]
        self.layout = LayoutManager(config)
        self.prs = Presentation()
        # 16:9 に固定
        self.prs.slide_width = self.layout.page_w
        self.prs.slide_height = self.layout.page_h

    def save(self, out_path: str):
        Path(out_path).parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(out_path)

    # ---------------- 内部ユーティリティ ----------------
    def _style_text(self, paragraph, text: str, size: Pt, bold=False,
                    italic=False, color=None, align=None):
        """統一的にフォントスタイルを適用"""
        paragraph.text = text
        run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
        run.font.name = self.fonts["family"]
        run.font.size = size
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color or self.colors["text"]
        if align:
            paragraph.alignment = align

    # ---------------- スライド実装 ----------------
    def add_title(self, data: Dict[str, Any]):
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])  # blank

        # タイトル
        rect = self.layout.get_rect("titleSlide.title")
        box = s.shapes.add_textbox(rect["left"], rect["top"], rect["width"], rect["height"])
        p = box.text_frame.paragraphs[0]
        self._style_text(
            p,
            data.get("title", ""),
            self.fonts["sizes"]["title"],
            bold=True,
            color=self.colors["primary"]
        )

        # 日付
        d_rect = self.layout.get_rect("titleSlide.date")
        dbox = s.shapes.add_textbox(d_rect["left"], d_rect["top"], d_rect["width"], d_rect["height"])
        dp = dbox.text_frame.paragraphs[0]
        self._style_text(
            dp,
            data.get("date", ""),
            self.fonts["sizes"]["body"],
            color=self.colors["subtext"]
        )
        return s

    def add_section(self, data: Dict[str, Any]):
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])  # blank

        # ゴースト番号
        g_rect = self.layout.get_rect("sectionSlide.ghostNum")
        gbox = s.shapes.add_textbox(g_rect["left"], g_rect["top"], g_rect["width"], g_rect["height"])
        gp = gbox.text_frame.paragraphs[0]
        self._style_text(
            gp,
            str(data.get("sectionNo", "01")),
            self.fonts["sizes"]["ghostNum"],
            bold=True,
            color=self.colors["ghost"]
        )

        # セクションタイトル
        t_rect = self.layout.get_rect("sectionSlide.title")
        tbox = s.shapes.add_textbox(t_rect["left"], t_rect["top"], t_rect["width"], t_rect["height"])
        tp = tbox.text_frame.paragraphs[0]
        self._style_text(
            tp,
            data.get("title", ""),
            self.fonts["sizes"]["sectionTitle"],
            bold=True,
            color=self.colors["text"],
            align=PP_ALIGN.CENTER
        )
        return s

    def add_content(self, data: Dict[str, Any]):
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])  # blank

        # Title
        t_rect = self.layout.get_rect("contentSlide.title")
        tbox = s.shapes.add_textbox(t_rect["left"], t_rect["top"], t_rect["width"], t_rect["height"])
        tp = tbox.text_frame.paragraphs[0]
        self._style_text(
            tp,
            data.get("title", ""),
            self.fonts["sizes"]["contentTitle"],
            bold=True,
            color=self.colors["primary"]
        )

        # Subhead
        subhead = data.get("subhead")
        if subhead:
            s_rect = self.layout.get_rect("contentSlide.subhead")
            sbox = s.shapes.add_textbox(s_rect["left"], s_rect["top"], s_rect["width"], s_rect["height"])
            sp = sbox.text_frame.paragraphs[0]
            self._style_text(
                sp,
                subhead,
                self.fonts["sizes"]["subhead"],
                color=self.colors["subtext"]
            )

        # Body (bullets)
        points: List[str] = data.get("points", [])
        b_rect = self.layout.get_rect("contentSlide.body")
        bbox = s.shapes.add_textbox(b_rect["left"], b_rect["top"], b_rect["width"], b_rect["height"])
        tf = bbox.text_frame
        tf.clear()
        for i, line in enumerate(points):
            p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
            self._style_text(
                p,
                line,
                self.fonts["sizes"]["body"],
                color=self.colors["text"]
            )

        return s
    
    def add_compare(self, data: Dict[str, Any]):
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])  # blank

        # タイトル
        t_rect = self.layout.get_rect("contentSlide.title")
        tbox = s.shapes.add_textbox(t_rect["left"], t_rect["top"], t_rect["width"], t_rect["height"])
        tp = tbox.text_frame.paragraphs[0]
        self._style_text(tp, data.get("title", "比較"),self.fonts["sizes"]["contentTitle"],bold=True, color=self.colors["primary"])

        # 左ボックス
        from pptx.util import Cm
        left_box = s.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Cm(1), Cm(4), Cm(12), Cm(8)
        )
        left_box.fill.solid()
        left_box.fill.fore_color.rgb = self.colors["surface"]
        left_tf = left_box.text_frame
        left_tf.text = data.get("leftTitle", "左側")
        for item in data.get("leftItems", []):
            p = left_tf.add_paragraph()
            self._style_text(p, f"• {item}", self.fonts["sizes"]["body"])

        # 右ボックス
        right_box = s.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Cm(14), Cm(4), Cm(12), Cm(8)
        )
        right_box.fill.solid()
        right_box.fill.fore_color.rgb = self.colors["surface"]
        right_tf = right_box.text_frame
        right_tf.text = data.get("rightTitle", "右側")
        for item in data.get("rightItems", []):
            p = right_tf.add_paragraph()
            self._style_text(p, f"• {item}", self.fonts["sizes"]["body"])

        return s

        # --- Cards ---
    def add_cards(self, data: Dict[str, Any]):
        """カード形式スライド"""
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])  # blank

        # タイトル
        t_rect = self.layout.get_rect("contentSlide.title")
        tbox = s.shapes.add_textbox(t_rect["left"], t_rect["top"], t_rect["width"], t_rect["height"])
        tp = tbox.text_frame.paragraphs[0]
        self._style_text(tp, data.get("title", "カード一覧"),self.fonts["sizes"]["contentTitle"], bold=True, color=self.colors["primary"])

        items = data.get("items", [])
        cols = min(3, max(1, int(data.get("columns", 3))))
        gap = Cm(0.5)
        card_w = (self.prs.slide_width - Cm(2) - gap * (cols - 1)) / cols
        rows = (len(items) + cols - 1) // cols
        card_h = Cm(5)

        for idx, item in enumerate(items):
            r, c = divmod(idx, cols)
            left = Cm(1) + c * (card_w + gap)
            top = Cm(4) + r * (card_h + gap)

            card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, card_w, card_h)
            card.fill.solid()
            card.fill.fore_color.rgb = self.colors["surface"]
            card.line.color.rgb = self.colors["ghost"]

            tf = card.text_frame
            tf.clear()

            if isinstance(item, dict):
                title = str(item.get("title", ""))
                desc = str(item.get("desc", ""))
                p = tf.paragraphs[0]
                self._style_text(p, title, self.fonts["sizes"]["body"], bold=True, color=self.colors["primary"])
                if desc:
                    p2 = tf.add_paragraph()
                    self._style_text(p2, desc, self.fonts["sizes"]["body"], color=self.colors["text"])
            else:
                p = tf.paragraphs[0]
                self._style_text(p, str(item), self.fonts["sizes"]["body"], color=self.colors["text"])

        return s
    
    # --- Progress ---
    def add_progress(self, data: Dict[str, Any]):
        """進捗バー"""
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])  # blank

        # タイトル
        t_rect = self.layout.get_rect("contentSlide.title")
        tbox = s.shapes.add_textbox(t_rect["left"], t_rect["top"], t_rect["width"], t_rect["height"])
        tp = tbox.text_frame.paragraphs[0]
        self._style_text(tp, data.get("title", "進捗状況"),
                     self.fonts["sizes"]["contentTitle"], bold=True, color=self.colors["primary"])

        items = data.get("items", [])
        bar_left = Cm(4)
        bar_width = Cm(18)
        bar_height = Cm(0.7)

        for i, item in enumerate(items):
            label = str(item.get("label", f"Step {i+1}"))
            pct = max(0, min(100, int(item.get("percent", 0))))

            # ラベル
            lbox = s.shapes.add_textbox(Cm(1), Cm(4 + i*2), Cm(3), Cm(1))
            lp = lbox.text_frame.paragraphs[0]
            self._style_text(lp, label, self.fonts["sizes"]["body"], color=self.colors["text"])

            # 背景バー
            bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, bar_left, Cm(4 + i*2), bar_width, bar_height)
            bg.fill.solid()
            bg.fill.fore_color.rgb = self.colors["ghost"]
            bg.line.fill.background()  # 枠なし

            # 実バー
            fg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, bar_left, Cm(4 + i*2), bar_width * pct / 100, bar_height)
            fg.fill.solid()
            fg.fill.fore_color.rgb = self.colors["accent"]
            fg.line.fill.background()

            # 数値ラベル
            pbox = s.shapes.add_textbox(bar_left + bar_width + Cm(0.3), Cm(4 + i*2), Cm(2), Cm(1))
            pp = pbox.text_frame.paragraphs[0]
            self._style_text(pp, f"{pct}%", self.fonts["sizes"]["body"], color=self.colors["subtext"])

        return s
    # --- Timeline ---
    def add_timeline(self, data: Dict[str, Any]):
        """タイムライン"""
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])  # blank

        # タイトル
        t_rect = self.layout.get_rect("contentSlide.title")
        tbox = s.shapes.add_textbox(t_rect["left"], t_rect["top"], t_rect["width"], t_rect["height"])
        tp = tbox.text_frame.paragraphs[0]
        self._style_text(tp, data.get("title", "タイムライン"),
                     self.fonts["sizes"]["contentTitle"], bold=True, color=self.colors["primary"])

        milestones = data.get("milestones", [])
        if not milestones:
            return s

        base_y = Cm(6)
        left_x = Cm(2)
        right_x = self.prs.slide_width - Cm(2)
        line = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_x, base_y, right_x - left_x, Cm(0.2))
        line.fill.solid()
        line.fill.fore_color.rgb = self.colors["ghost"]
        line.line.fill.background()

        gap = (right_x - left_x) / max(1, len(milestones)-1)
        dot_r = Cm(0.5)

        for i, m in enumerate(milestones):
            x = left_x + gap * i - dot_r/2
            dot = s.shapes.add_shape(MSO_SHAPE.OVAL, x, base_y - dot_r/2, dot_r, dot_r)
            dot.fill.solid()
            dot.fill.fore_color.rgb = self.colors["accent"]
            dot.line.fill.background()

            # ラベル
            lbox = s.shapes.add_textbox(x - Cm(1), base_y - Cm(1.5), Cm(3), Cm(0.7))
            lp = lbox.text_frame.paragraphs[0]
            self._style_text(lp, str(m.get("label", "")),
                         self.fonts["sizes"]["body"], bold=True, color=self.colors["text"],
                         align=PP_ALIGN.CENTER)

            # 日付
            dbox = s.shapes.add_textbox(x - Cm(1), base_y + Cm(0.5), Cm(3), Cm(0.7))
            dp = dbox.text_frame.paragraphs[0]
            self._style_text(dp, str(m.get("date", "")),
                         self.fonts["sizes"]["caption"], color=self.colors["subtext"],
                         align=PP_ALIGN.CENTER)

        return s

    # ---------------------- ビルド関数 ----------------------
def build_pptx_from_plan(plan: Dict[str, Any], out_path: str):
    sf = SlideFactory()
    for spec in plan.get("slides", []):
        t = spec.get("type")
        if t=="title": sf.add_title(spec)
        elif t=="section": sf.add_section(spec)
        elif t=="content": sf.add_content(spec)
        elif t=="compare": sf.add_compare(spec)
        elif t=="cards": sf.add_cards(spec)
        elif t=="progress": sf.add_progress(spec)
        elif t=="timeline": sf.add_timeline(spec)
    sf.save(out_path)



# ---------------- CLI ----------------
if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python make_slides.py plan.json out.pptx")
        sys.exit(1)

    plan_path = Path(sys.argv[1])
    out_path = sys.argv[2]
    with plan_path.open("r", encoding="utf-8") as f:
        plan = json.load(f)

    build_pptx_from_plan(plan, out_path)
    print(f"✅ Done: {out_path}")
