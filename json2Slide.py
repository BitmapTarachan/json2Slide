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
from pptx.enum.text import MSO_ANCHOR


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
            "contentTitle": Pt(30),
            "subhead": Pt(28),
            "body": Pt(22),
            "caption": Pt(18),
            "ghostNum": Pt(180),
        }
    },

    # ---------------- レイアウト座標 ----------------
    "POS_PX": {
        "titleSlide": {
            "subject":  {"left": 50, "top": 140, "width": 800, "height": 40},  
                "title":    {"left": 50, "top": 190, "width": 800, "height": 90},  
                "lecturer": {"left": 50, "top": 290, "width": 400, "height": 40},  
                "date":     {"left": 50, "top": 330, "width": 250, "height": 40},  
        },        
        "sectionSlide": {
            "title": {"left": 55, "top": 230, "width": 840, "height": 80},
            "ghostNum": {"left": 100, "top": 120, "width": 300, "height": 200},
        },
        "contentSlide": {
            "title": {"left": 25, "top": 40, "width": 830, "height": 50},
            "subhead": {"left": 25, "top": 100, "width": 830, "height": 30},
            "body": {"left": 25, "top": 150, "width": 910, "height": 303},
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

        # 教科名
        subj_rect = self.layout.get_rect("titleSlide.subject")
        subj_box = s.shapes.add_textbox(subj_rect["left"], subj_rect["top"], subj_rect["width"], subj_rect["height"])
        subj_p = subj_box.text_frame.paragraphs[0]
        self._style_text(
            subj_p,
            data.get("subject", ""),
            self.fonts["sizes"]["contentTitle"],
            color=self.colors["subtext"]
        )

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

        # 講師名（固定）
        l_rect = self.layout.get_rect("titleSlide.lecturer")
        lbox = s.shapes.add_textbox(l_rect["left"], l_rect["top"], l_rect["width"], l_rect["height"])
        lp = lbox.text_frame.paragraphs[0]
        self._style_text(
            lp,
            "講師名：〇〇　〇〇",
            self.fonts["sizes"]["body"],
            color=self.colors["text"]
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
            sbox.text_frame.word_wrap = True

        # Body area
        b_rect = self.layout.get_rect("contentSlide.body")

        points: List[str] = data.get("points", [])
        body_text: str = data.get("bodyText", "")

        last_y = b_rect["top"]

        # --- 箇条書き ---
        if points:
            # 箇条書き部分
            line_spacing = self.fonts["sizes"]["body"].pt + 10  # フォントサイズ + 行間
            bbox = s.shapes.add_textbox(b_rect["left"], b_rect["top"], b_rect["width"], b_rect["height"])
            tf = bbox.text_frame
            tf.clear()
            tf.word_wrap = True

            for i, line in enumerate(points):
                p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
                self._style_text(p, line, self.fonts["sizes"]["body"], color=self.colors["text"])
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
            self._style_text(lp, body_text, self.fonts["sizes"]["body"], color=self.colors["text"])

        return s
    
    # 比較    
    def add_compare(self, data: Dict[str, Any]):
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])  # blank

        # スライドタイトル
        t_rect = self.layout.get_rect("contentSlide.title")
        tbox = s.shapes.add_textbox(t_rect["left"], t_rect["top"], t_rect["width"], t_rect["height"])
        tp = tbox.text_frame.paragraphs[0]
        self._style_text(tp, data.get("title", "比較"),
                         self.fonts["sizes"]["contentTitle"], bold=True, color=self.colors["primary"])

        # ボックス配置
        margin = Cm(1.5)
        gap = Cm(1.5)
        box_w = (self.prs.slide_width - margin * 2 - gap) / 2
        box_h = Cm(8)
        top = Cm(4)

        def add_box(x, title, items):
            box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, top, box_w, box_h)
            box.fill.solid()
            box.fill.fore_color.rgb = self.colors["surface"]
            box.line.color.rgb = self.colors["ghost"]

            tf = box.text_frame
            tf.clear()
            tf.word_wrap = True

            # タイトル
            p = tf.paragraphs[0]
            self._style_text(p, title, self.fonts["sizes"]["subhead"], bold=True, color=self.colors["text"])
            p.space_after = Pt(15)

            # 箇条書き
            for item in items:
                para = tf.add_paragraph()
                self._style_text(para, f"• {item}", self.fonts["sizes"]["body"], color=self.colors["text"])
                para.space_after = Pt(6)

        # 左ボックス
        add_box(margin,
                data.get("leftTitle", "選択肢A"),
                data.get("leftItems", [])
        )

        # 右ボックス
        add_box(margin + box_w + gap,
                data.get("rightTitle", "選択肢B"),
                data.get("rightItems", [])
        )

        # --- 結論 BodyText ---
        body_text = data.get("bodyText", "")
        if body_text:
            b_rect = self.layout.get_rect("contentSlide.body")
            lbox = s.shapes.add_textbox(b_rect["left"], top + box_h + Cm(1.0), b_rect["width"], Cm(3))
            tf2 = lbox.text_frame
            tf2.word_wrap = True
            p = tf2.paragraphs[0]
            self._style_text(p, body_text, self.fonts["sizes"]["body"], color=self.colors["text"], align=PP_ALIGN.CENTER)

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
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.TOP  

            if isinstance(item, dict):
                title = str(item.get("title", ""))
                desc = str(item.get("desc", ""))
                p = tf.paragraphs[0]
                self._style_text(p, title, self.fonts["sizes"]["body"], bold=True, color=self.colors["primary"])
                if desc:
                    p2 = tf.add_paragraph()
                    p2.space_before = Pt(8) 
                    self._style_text(p2, desc, self.fonts["sizes"]["caption"], color=self.colors["text"])
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
        bar_left = Cm(6)
        bar_width = Cm(23) 
        bar_height = Cm(1)
        v_gap = Cm(1.5)

        for i, item in enumerate(items):
            label = str(item.get("label", f"Step {i+1}"))
            pct = max(0, min(100, int(item.get("percent", 0))))

            y = Cm(4) + i * v_gap

            # ラベル
            lbox = s.shapes.add_textbox(Cm(0.5), y - Cm(0.2), bar_left - Cm(1), bar_height)
            lp = lbox.text_frame.paragraphs[0]
            self._style_text(lp, label, self.fonts["sizes"]["body"], color=self.colors["text"], align=PP_ALIGN.RIGHT)

            # 背景バー
            bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, bar_left, y, bar_width, bar_height)
            bg.fill.solid()
            bg.fill.fore_color.rgb = self.colors["ghost"]
            bg.line.fill.background()

            # 実バー
            fg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, bar_left, y, bar_width * pct / 100, bar_height)
            fg.fill.solid()
            fg.fill.fore_color.rgb = self.colors["accent"]
            fg.line.fill.background()

            # 数値ラベル（右側も余裕広く）
            pbox = s.shapes.add_textbox(bar_left + bar_width + Cm(0.5), y - Cm(0.2), Cm(4), bar_height)
            pp = pbox.text_frame.paragraphs[0]
            self._style_text(pp, f"{pct}%", self.fonts["sizes"]["body"], color=self.colors["subtext"])

        return s
    
    def add_timeline(self, data: Dict[str, Any]):
        """タイムライン"""
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])  # blank

        # タイトル
        t_rect = self.layout.get_rect("contentSlide.title")
        tbox = s.shapes.add_textbox(t_rect["left"], t_rect["top"], t_rect["width"], t_rect["height"])
        tp = tbox.text_frame.paragraphs[0]
        self._style_text(tp, data.get("title", "タイムライン"),
                         self.fonts["sizes"]["contentTitle"], bold=True, color=self.colors["primary"])


        slide_height = self.prs.slide_height

        # バー位置をスライド中央に配置
        bar_height = Cm(0.25)
        bar_top = (slide_height - bar_height) / 2   # 上下中央
        bar_left = Cm(4)                            # ← 左右を少し余白大きめに
        bar_width = self.prs.slide_width - Cm(8)    # ← margin=4cmずつ確保

        # ベースバー
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, bar_left, bar_top, bar_width, bar_height)
        bar.fill.solid()
        bar.fill.fore_color.rgb = self.colors["ghost"]
        bar.line.fill.background()

        milestones = data.get("milestones", [])
        if not milestones:
            return s  # データがなければラインだけで終了

        step_x = bar_width / (len(milestones)-1 if len(milestones) > 1 else 1)

        for i, m in enumerate(milestones):
            title = str(m.get("label", f"Step {i+1}"))
            date = str(m.get("date", ""))

            x = bar_left + i * step_x

            # マーカー
            circle = s.shapes.add_shape(MSO_SHAPE.OVAL, x-Cm(0.2), bar_top-Cm(0.2), Cm(0.4), Cm(0.4))
            circle.fill.solid()
            circle.fill.fore_color.rgb = self.colors["accent"]
            circle.line.fill.background()

            # タイトル（2行分確保・下揃え）
            tbox = s.shapes.add_textbox(x-Cm(2), bar_top-Cm(2), Cm(4), Cm(1.5))
            tf = tbox.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.BOTTOM
            p = tf.paragraphs[0]
            self._style_text(p, title, self.fonts["sizes"]["body"], bold=True, color=self.colors["text"], align=PP_ALIGN.CENTER)

            # 日付
            dbox = s.shapes.add_textbox(x-Cm(2), bar_top+Cm(0.5), Cm(4), Cm(1))
            dp = dbox.text_frame.paragraphs[0]
            self._style_text(dp, date, self.fonts["sizes"]["caption"], color=self.colors["subtext"], align=PP_ALIGN.CENTER)
    
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
        print("Usage: python json2.py plan.json out.pptx")
        sys.exit(1)

    plan_path = Path(sys.argv[1])
    out_path = sys.argv[2]
    with plan_path.open("r", encoding="utf-8") as f:
        plan = json.load(f)

    build_pptx_from_plan(plan, out_path)
    print(f"✅ Done: {out_path}")
