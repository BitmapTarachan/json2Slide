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
import requests
import io
import math


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
from pptx.oxml.xmlchemy import OxmlElement

from PIL import Image


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
        "primary"    : RGBColor(0x42, 0x85, 0xF4),   # メインカラー（タイトル・強調）
        "accent"     : RGBColor(0xFB, 0xBC, 0x04),   # アクセントカラー（ハイライト・バー）
        "background" : RGBColor(0xFF, 0xFF, 0xFF),   # スライド背景
        "surface"    : RGBColor(0xF8, 0xF9, 0xFA),   # セクション背景・ボックス背景
        "text"       : RGBColor(0x33, 0x33, 0x33),   # 標準本文
        "subtext"    : RGBColor(0x9E, 0x9E, 0x9E),   # 補助テキスト（日付・キャプション）
        "ghost"      : RGBColor(0xEF, 0xEF, 0xED),   # ゴースト数字・区切り用
    },

    # ---------------- フォント定義 ----------------
    "FONTS": {
        "family": "Biz UDゴシック",
        "sizes": {
            "title"        : Pt(45),
            "sectionTitle" : Pt(38),
            "contentTitle" : Pt(30),
            "subhead"      : Pt(28),
            "body"         : Pt(22),
            "caption"      : Pt(18),
            "ghostNum"     : Pt(180),
        }
    },

    # ---------------- レイアウト座標 ----------------
    "POS_PX": {
        "titleSlide": {
            "subject":  { "left": 80, "top": 140, "width": 800, "height": 40},  
                "title"    : { "left": 80, "top": 190, "width": 800, "height": 90 },  
                "lecturer" : { "left": 80, "top": 290, "width": 400, "height": 40 },  
                "date"     : { "left": 80, "top": 330, "width": 250, "height": 40 },  
        },        
        "sectionSlide": {
            "title"    : { "left":  55, "top": 230, "width": 840, "height":  80 },
            "ghostNum" : { "left": 100, "top": 120, "width": 300, "height": 200 },
        },
        "contentSlide": {
            "title"    : { "left": 28, "top":  30, "width": 830, "height":  50 },
            "subhead"  : { "left": 25, "top": 100, "width": 830, "height":  30 },
            "body"     : { "left": 25, "top": 150, "width": 910, "height": 303 },
        }
    }
}

THEMES = {
    "Default": {
        "primary"    : RGBColor(0x42, 0x85, 0xF4),  # Googleブルー
        "accent"     : RGBColor(0xFB, 0xBC, 0x04),  # 黄色
        "background" : RGBColor(0xFF, 0xFF, 0xFF),
        "surface"    : RGBColor(0xF8, 0xF9, 0xFA),
        "text"       : RGBColor(0x33, 0x33, 0x33),
        "subtext"    : RGBColor(0x9E, 0x9E, 0x9E),
        "ghost"      : RGBColor(0xEF, 0xEF, 0xED),
    },
    "Nature": {
        "primary"    : RGBColor(0x2E, 0x7D, 0x32),  # 深緑
        "accent"     : RGBColor(0xFF, 0xA0, 0x00),  # オレンジ
        "background" : RGBColor(0xFF, 0xFF, 0xF5),
        "surface"    : RGBColor(0xE8, 0xF5, 0xE9),
        "text"       : RGBColor(0x1B, 0x5E, 0x20),
        "subtext"    : RGBColor(0x6D, 0x6D, 0x6D),
        "ghost"      : RGBColor(0xC8, 0xE6, 0xC9),
    },
    "Dark": {
        "primary"    : RGBColor(0xBB, 0x86, 0xFC),  # 紫
        "accent"     : RGBColor(0x03, 0xDA, 0xC6),  # シアン
        "background" : RGBColor(0x12, 0x12, 0x12),
        "surface"    : RGBColor(0x1E, 0x1E, 0x1E),
        "text"       : RGBColor(0xEE, 0xEE, 0xEE),
        "subtext"    : RGBColor(0xAA, 0xAA, 0xAA),
        "ghost"      : RGBColor(0x33, 0x33, 0x33),
    },
    "Monochrome": {
        "primary"    : RGBColor(0x00, 0x00, 0x00),
        "accent"     : RGBColor(0x55, 0x55, 0x55),
        "background" : RGBColor(0xFF, 0xFF, 0xFF),
        "surface"    : RGBColor(0xF0, 0xF0, 0xF0),
        "text"       : RGBColor(0x00, 0x00, 0x00),
        "subtext"    : RGBColor(0x77, 0x77, 0x77),
        "ghost"      : RGBColor(0xDD, 0xDD, 0xDD),
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

def hex_to_rgbcolor(hex_str: str) -> RGBColor:
    hex_str = hex_str.lstrip("#")
    r, g, b = int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16)
    return RGBColor(r, g, b)

def ensure_list(x) -> List[Any]:
    if x is None:
        return []
    return x if isinstance(x, list) else [x]


# ---------------- Slide Factory ----------------
class SlideFactory:
    def __init__(self, plan: Dict[str, Any], config=CONFIG):
        # イメージキャッシュ
        self._image_cache = {}

        # カラーテーマ選択S
        theme_name = plan.get("color-theme", "Default")

        if theme_name == "Custom" and "colors" in plan:
            self.colors = {
                k: hex_to_rgbcolor(v) for k, v in plan["colors"].items()
            }
        else:
            self.colors = THEMES.get(theme_name, THEMES["Default"])

        self.config = config
        self.fonts = config["FONTS"]
        self.layout = LayoutManager(config)
        self.prs = Presentation()
        # 16:9 に固定
        self.prs.slide_width = self.layout.page_w
        self.prs.slide_height = self.layout.page_h
        
        # 全体背景
        self._global_bg = None
        if plan.get("background-image"):
            self._global_bg, _ = self._load_image(plan["background-image"])

    def save(self, out_path: str):
        Path(out_path).parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(out_path)

    # ---------------- 内部ユーティリティ ----------------
    def _new_slide(self, data: Dict[str, Any],apply_background = True):
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        # 背景
        self._apply_background(s,data,apply_background)
        # スライドノート
        note_text = data.get("note", "")
        s.notes_slide.notes_text_frame.text = note_text
        return s

    def _style_text(self, paragraph, text: str, size: Pt, bold=False,
                    italic=False, color=None, align=None):
        """統一的にフォントスタイルを適用"""
        paragraph.text = text
        run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
        run.font.name = self.fonts["family"]
        run.font.size = size
        run.font.bold = bool(bold) if bold is not None else None
        run.font.italic = italic
        run.font.color.rgb = color or self.colors["text"]
        if align:
            paragraph.alignment = align

    def _load_image(self, path_or_url: str):
        try:
            # キャッシュヒット
            if path_or_url in self._image_cache:
                # BytesIOは再利用のため毎回seek(0)して返す
                stream, im = self._image_cache[path_or_url]
                stream.seek(0)
                return stream, im

            # URLから取得
            if path_or_url.startswith(("http://", "https://")):
                response = requests.get(path_or_url, timeout=10)
                response.raise_for_status()
                stream = io.BytesIO(response.content)
            else:
                # ローカルファイル
                with open(path_or_url, "rb") as f:
                    stream = io.BytesIO(f.read())

            # PILでロードしてキャッシュ
            im = Image.open(stream)
            im.load()           # Lazy読み込みを強制
            stream.seek(0)      # 再利用に備えて戻す
            self._image_cache[path_or_url] = (stream, im)

            return stream, im
        except Exception as e:
            print(f"[WARN] 画像を読み込めませんでした: {path_or_url} ({e})")
            return None, None
    
    def _add_slide_title(self, slide, title: str):
        """
        スライドタイトルを描画する共通関数
        左端にアクセントカラーの縦長バーを置き、
        その右にタイトルテキストを配置する。
        """
        # レイアウトからタイトル領域を取得
        t_rect = self.layout.get_rect("contentSlide.title")
        left, top, width, height = t_rect["left"], t_rect["top"], t_rect["width"], t_rect["height"]

        # --- 縦長バー ---
        bar_width = Pt(6)   # 適度に細いバー
        bar_margin = Pt(4)  # テキストとの間隔

        slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left, top,
            bar_width, height
        ).fill.solid()
        slide.shapes[-1].fill.fore_color.rgb = self.colors["accent"]
        slide.shapes[-1].line.fill.background()  # 枠線なし
        slide.shapes[-1].shadow.inherit = False  # 影を消す（フラット）

        # --- タイトルテキスト ---
        tbox = slide.shapes.add_textbox(
            left + bar_width + bar_margin, top,
            width - bar_width - bar_margin, height
        )
        tf = tbox.text_frame
        tp = tf.paragraphs[0]
        self._style_text(
            tp,
            title,
            self.fonts["sizes"]["contentTitle"],
            bold=True,
            color=self.colors["primary"]
        )
        tp.alignment = PP_ALIGN.LEFT
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    def _apply_background(self, slide, slide_data: dict = None,apply_background = True):
        bg_url = None
        if slide_data and "background-image" in slide_data:
            bg_url = slide_data["background-image"]
            stream, _ = self._load_image(bg_url)
            if stream and _:
                slide.shapes.add_picture(stream, 0, 0,
                                        width=self.prs.slide_width,
                                        height=self.prs.slide_height)
        elif apply_background and self._global_bg:
            if self._global_bg :
                slide.shapes.add_picture(self._global_bg, 0, 0,
                                        width=self.prs.slide_width,
                                        height=self.prs.slide_height)
        else:
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = self.colors["background"]    
            
    def _set_shape_transparency(self, shape, alpha_val: int):
        """
        shape.fill に alpha 要素を追加して透過を指定（alpha_val は 0〜100000）
        例: 40000 = 40% 透過
        """
        from pptx.oxml.xmlchemy import OxmlElement
        ts = shape.fill._xPr.solidFill
        sF = ts.get_or_change_to_srgbClr()
        alpha_elem = OxmlElement('a:alpha')
        alpha_elem.set('val', str(alpha_val))
        sF.append(alpha_elem)

    # ---------------- スライド実装 ----------------
    def add_title(self, data: Dict[str, Any]):
        s = self._new_slide(data)
        slide_w, slide_h = self.prs.slide_width, self.prs.slide_height

        # 右側に画像を配置（オプション）
        if "image" in data and data["image"]:
            stream, im = self._load_image(data["image"])
            if stream and im:
                iw, ih = im.size
                slide_w, slide_h = self.prs.slide_width, self.prs.slide_height

                # 縦に合わせる（はみ出さないように調整）
                scale = slide_h / ih
                new_w, new_h = iw * scale, ih * scale

                left = int(slide_w * 3/5)
                top = 0
                if left + new_w < slide_w:  # 幅足りないなら右寄せ
                    left = slide_w - new_w

                s.shapes.add_picture(stream, left, top, width=int(new_w), height=int(new_h))

        # 左側の縦線（Ghostカラー）
        line_left = Pt(50)  # 余白を少し空ける
        s.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            line_left, Pt(90),  # X, Y
            Pt(2), slide_h - Pt(200)  
        ).fill.solid()
        s.shapes[-1].fill.fore_color.rgb = self.colors["ghost"]
        s.shapes[-1].line.fill.background()  # 枠線なし
        s.shapes[-1].shadow.inherit = False  # 影無し

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
        s = self._new_slide(data)

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
        s = self._new_slide(data)
        self._add_slide_title(s, data.get("title",""))

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
        s = self._new_slide(data)
        self._add_slide_title(s, data.get("title","比較"))

        # ボックス配置
        margin = Cm(1.5)
        gap = Cm(1.5)
        box_w = (self.prs.slide_width - margin * 2 - gap) / 2
        box_h = Cm(8)
        top = Cm(4)

        def add_box(x, title, items):
            box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, top, box_w, box_h)
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
        s = self._new_slide(data)
        self._add_slide_title(s, data.get("title","一覧"))

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
        s = self._new_slide(data)

        self._add_slide_title(s, data.get("title","進捗状況"))

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
        s = self._new_slide(data)
        self._add_slide_title(s, data.get("title",""))

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

    def add_image_auto(self, data: Dict[str, Any]):
        images = data.get("images", [])
        n = len(images)
        if n == 0:
            return

        # スライド作成
        slide = self._new_slide(data)

        # 画像レイアウト分岐
        if n == 1:
            self._add_image_rightcontent(slide, images, Pt(24))
        elif n == 2:
            self._add_image_twocol(slide, images, Pt(22))
        elif n == 3:
            self._add_image_three_grid(slide, images, Pt(18))
        elif n == 4:
            self._add_image_four_grid(slide, images, Pt(18))

        return slide    # --- 1枚: 右にキャプション ---
    
    # --- 1枚: 画像 + 右にキャプション ---
    def _add_image_rightcontent(self, slide, images, font_size):
        img = images[0]
        caption = img.get("caption", "")

        # スライドサイズ
        slide_width = self.prs.slide_width
        slide_height = self.prs.slide_height
        margin = Pt(20)

        # self._load_image で画像読み込み（BytesIO, PIL.Image）
        stream, im = self._load_image(img["url"])
        if stream :
            iw, ih = im.size
            aspect = iw / ih

            # 横長か縦長かでリサイズ基準を切替
            if aspect >= 1:  # 横長 → 横幅を中央まで広げる
                max_width = slide_width / 2 - 2 * margin
                width = max_width
                height = width / aspect
                left = margin
                top = (slide_height - height) / 2
            else:  # 縦長 → 上下いっぱいまで
                max_height = slide_height - 2 * margin
                height = max_height
                width = height * aspect
                left = margin
                top = (slide_height - height) / 2

            # 画像挿入
            slide.shapes.add_picture(stream, left, top, width=width, height=height)

        # キャプションテキスト
        cap_left = slide_width / 2 + margin
        cap_width = slide_width / 2 - 2 * margin
        cap_height = slide_height - 2 * margin
        cap_top = margin

        txBox = slide.shapes.add_textbox(cap_left, cap_top, cap_width, cap_height)
        tf = txBox.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = caption
        p.alignment = PP_ALIGN.LEFT
        self._style_text(p, caption, font_size, self.colors["text"])


    # --- 2枚: 縦に並べて右にキャプション ---
    def _add_image_twocol(self, slide, images, font_size):
        slide_w, slide_h = self.prs.slide_width, self.prs.slide_height
        margin = Pt(40)
        spacing = Pt(30)
        title_height = Pt(60)  # タイトル領域を固定値で確保

        # 画像の最大表示高さ
        max_h = (slide_h - margin*2 - spacing - title_height) / 2
        target_w = slide_w / 2 - margin * 2

        # 画像サイズを読み取り、リサイズ結果を保存
        scaled_sizes = []
        streams = []
        for img in images:
            stream, im = self._load_image(img["url"])
            if stream:
                iw, ih = im.size
                scale = min(max_h/ih, target_w/iw)  # 高さ基準 + 横幅制限
                new_w, new_h = iw*scale, ih*scale
                scaled_sizes.append((new_w, new_h))
                streams.append(stream)

        # 上下サイズをそろえる（小さい方に合わせる）
        min_h = min(h for _, h in scaled_sizes)
        scaled_sizes = [(w*(min_h/h), min_h) for w, h in scaled_sizes]

        # 全体の高さ（上下 + spacing）
        total_h = scaled_sizes[0][1] + scaled_sizes[1][1] + spacing
        top_start = (slide_h - total_h) / 2 + title_height/2

        shapes = []
        for i, ((new_w, new_h), stream, img) in enumerate(zip(scaled_sizes, streams, images)):
            top = top_start + i*(new_h + spacing)
            left = (slide_w/2 - new_w) / 2

            # 画像配置
            if stream:
                pic = slide.shapes.add_picture(stream, left, top, width=int(new_w), height=int(new_h))  

            # キャプションを画像の右に配置
            cap_box = slide.shapes.add_textbox(
                left+new_w+Pt(20), top, slide_w/2 - new_w - Pt(40), new_h
            )
            tf = cap_box.text_frame
            tf.word.wrap = True
            p = tf.paragraphs[0]
            self._style_text(p, img.get("caption",""), font_size, self.colors["text"])
            p.alignment = PP_ALIGN.LEFT
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            shapes.append((pic, cap_box))

        return shapes

        # --- 3枚: 横3グリッド + 下キャプション ---
    def _add_image_three_grid(self, slide, images, font_size):
        slide_w, slide_h = self.prs.slide_width, self.prs.slide_height
        margin = Pt(40)
        spacing = Pt(30)
        max_width_ratio = 0.8  # 横幅全体の80%に収める

        # 各画像の幅（等分）
        target_w = (slide_w * max_width_ratio - 2 * spacing) / 3

        scaled_sizes = []
        streams = []
        for img in images:
            stream, im = self._load_image(img["url"])
            if stream :
                iw, ih = im.size
                scale = target_w / iw
                new_w, new_h = iw * scale, ih * scale
                scaled_sizes.append((int(new_w), int(new_h)))
                streams.append(stream)

        # 横方向の開始位置（中央寄せ）
        total_w = sum(w for w, h in scaled_sizes) + 2 * spacing
        left_start = (slide_w - total_w) / 2

        # 最大高さを揃える（縦位置は上を揃える）
        img_max_h = max(h for w, h in scaled_sizes)
        top_img = (slide_h / 2) - img_max_h / 2 - Pt(20)

        # キャプションのY座標（全画像共通で揃える）
        cap_top = top_img + img_max_h + Pt(10)

        shapes = []
        x = left_start
        for (w, h), stream, img in zip(scaled_sizes, streams, images):
            # 画像（上を揃える。小さい画像は下に余白）
            top = top_img + (img_max_h - h)
            if stream:
                slide.shapes.add_picture(stream, x, top, width=w, height=h)

            # キャプション（Y位置を固定）
            cap_box = slide.shapes.add_textbox(x, cap_top, w, Pt(40))
            tf = cap_box.text_frame
            tf.word.wrap = True
            p = tf.paragraphs[0]
            self._style_text(p, img.get("caption", ""), font_size, self.colors["text"])
            p.alignment = PP_ALIGN.LEFT
            tf.vertical_anchor = MSO_ANCHOR.TOP
            shapes.append(cap_box)

            x += w + spacing

        return shapes

    # --- 4枚: 横4グリッド + 下キャプション ---
    def _add_image_four_grid(self, slide, images, font_size):
        slide_w, slide_h = self.prs.slide_width, self.prs.slide_height
        margin = Pt(40)
        spacing = Pt(30)
        n = 4

        # 横方向の基準幅（スライド幅の90%に収める）
        max_total_w = slide_w * 0.9
        target_w = (max_total_w - spacing * (n - 1)) / n

        resized = []
        streams = []

        for img in images:
            stream, im = self._load_image(img["url"])
            if stream:
                iw, ih = im.size
                scale = target_w / iw
                new_w, new_h = iw * scale, ih * scale
                resized.append((new_w, new_h, img))
                streams.append(stream)

        # 最大高さを取得（キャプション基準にする）
        max_h = max(h for _, h, _ in resized)

        # 横方向の開始位置（中央寄せ）
        total_w = sum(w for w, _, _ in resized) + spacing * (n - 1)
        left_base = (slide_w - total_w) / 2
        top_img = slide_h * 0.35

        shapes = []
        cur_left = left_base
        for (w, h, img), stream in zip(resized, streams):
            # 画像（下揃え）
            pic_top = top_img + (max_h - h)
            if stream: 
                pic = slide.shapes.add_picture(stream, cur_left, pic_top, width=int(w), height=int(h))

            # キャプション（全画像上揃え）
            caption = img.get("caption", "")
            cap_box = slide.shapes.add_textbox(cur_left, top_img + max_h + Pt(10), w, Pt(40))
            tf = cap_box.text_frame
            p = tf.paragraphs[0]
            self._style_text(p, caption, font_size, color=self.colors["text"])
            p.alignment = PP_ALIGN.LEFT

            if stream:
                shapes.append((pic, cap_box))

            cur_left += w + spacing

        return shapes

    # --- Q&A: Question ---
    def add_qa_question(self, data: Dict[str, Any]):
        s = self._new_slide(data)
        slide_w, slide_h = self.prs.slide_width, self.prs.slide_height

        # ゴースト "Ｑ"（全角、大きく左上）
        q_size = int(min(slide_w, slide_h) * 0.5)  # スライドの半分くらいの大きさ
        qbox = s.shapes.add_textbox(Pt(0), Pt(0), q_size, q_size)
        tf_q = qbox.text_frame
        tf_q.word_wrap = False
        tf_q.vertical_anchor = MSO_ANCHOR.TOP

        qp = tf_q.paragraphs[0]
        run = qp.add_run()
        run.text = "Ｑ"
        run.font.size = Pt(200)   # ゴーストQ専用サイズ（必要に応じて調整）
        run.font.bold = True
        run.font.color.rgb = self.colors["ghost"]
        qp.alignment = PP_ALIGN.LEFT

        # 質問文（中央揃え）
        qtext_box = s.shapes.add_textbox(Pt(100), slide_h/3, slide_w - Pt(200), slide_h/3)
        tf = qtext_box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        self._style_text(
            p,
            data.get("question", ""),
            self.fonts["sizes"]["sectionTitle"],
            bold=True,
            color=self.colors["text"],
            align=PP_ALIGN.CENTER
        )

        return s

    # --- Q&A: Answer ---
    def add_qa_answer(self, data: Dict[str, Any]):
        s = self._new_slide(data)
        slide_w, slide_h = self.prs.slide_width, self.prs.slide_height

        # ゴースト "Ａ"（全角、大きく左上）
        a_size = int(min(slide_w, slide_h) * 0.5)
        abox = s.shapes.add_textbox(Pt(0), Pt(0), a_size, a_size)
        tf_a = abox.text_frame
        tf_a.word_wrap = False
        tf_a.vertical_anchor = MSO_ANCHOR.TOP

        ap = tf_a.paragraphs[0]
        run = ap.add_run()
        run.text = "Ａ"
        run.font.size = Pt(200)   # ゴーストA専用サイズ
        run.font.bold = True
        run.font.color.rgb = self.colors["ghost"]
        ap.alignment = PP_ALIGN.LEFT

        # 答え（中央に一言）
        ans_box = s.shapes.add_textbox(Pt(100), Pt(200), slide_w - Pt(200), Pt(100))
        tf_ans = ans_box.text_frame
        ans_p = tf_ans.paragraphs[0]
        self._style_text(
            ans_p,
            "答え : " + data.get("answer", ""),
            self.fonts["sizes"]["contentTitle"],
            bold=True,
            color=self.colors["primary"],
            align=PP_ALIGN.CENTER
        )

        # 解説（中央寄せ）
        exp_box = s.shapes.add_textbox(Pt(100), slide_h/3, slide_w - Pt(200), slide_h/2)
        tf_exp = exp_box.text_frame
        tf_exp.word_wrap = True
        tf_exp.vertical_anchor = MSO_ANCHOR.MIDDLE
        exp_p = tf_exp.paragraphs[0]
        self._style_text(
            exp_p,
            data.get("explanation", ""),
            self.fonts["sizes"]["body"],
            color=self.colors["text"],
            align=PP_ALIGN.CENTER
        )

        return s
    
    # 表形式
    def add_table_slide(self, data: Dict[str, Any]):
        s = self._new_slide(data)
        slide_w, slide_h = self.prs.slide_width, self.prs.slide_height

        self._add_slide_title(s, data.get("title", "表"))

        headers = data.get("headers", [])
        rows = data.get("rows", [])
        n_rows, n_cols = len(rows) + 1, len(headers)

        top = Pt(100)
        left = Pt(40)
        width = int(slide_w - Pt(80))
        height = int(slide_h * 0.55)

        table_shape = s.shapes.add_table(n_rows, n_cols, left, top, width, height)
        table = table_shape.table

        # 列幅（整数化必須）
        col_width = int(width / n_cols)
        for col in table.columns:
            col.width = col_width

        # 行高さ（整数化必須）
        row_height = int(height / n_rows)
        for row in table.rows:
            row.height = row_height

        # ヘッダー
        for j, header in enumerate(headers):
            cell = table.cell(0, j)
            cell.text_frame.clear()
            p = cell.text_frame.paragraphs[0]
            run = p.add_run()
            run.text = header
            run.font.size = Pt(18)
            run.font.bold = True
            run.font.name = "BIZ UDゴシック"
            run.font.color.rgb = self.colors["background"]
            cell.fill.solid()
            cell.fill.fore_color.rgb = self.colors["primary"]
            p.alignment = PP_ALIGN.CENTER

        # データ
        for i, row_data in enumerate(rows):
            for j, val in enumerate(row_data):
                cell = table.cell(i + 1, j)
                cell.text_frame.clear()
                p = cell.text_frame.paragraphs[0]
                run = p.add_run()
                run.text = str(val)
                run.font.size = Pt(16)
                run.font.name = "BIZ UDゴシック"
                run.font.color.rgb = self.colors["text"]
                p.alignment = PP_ALIGN.CENTER
                if i % 2 == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = self.colors["surface"]
                else:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = self.colors["background"]

        # bodyText（高さが溢れないように制限）
        body_text = data.get("bodyText")
        if body_text:
            b_top = min(top + height + Pt(20), slide_h - Pt(100))
            b_left = Pt(40)
            b_width = int(slide_w - Pt(80))
            b_height = int(slide_h - b_top - Pt(40))

            box = s.shapes.add_textbox(b_left, b_top, b_width, b_height)
            tf = box.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = body_text
            run.font.size = Pt(16)
            run.font.name = "BIZ UDゴシック"
            run.font.color.rgb = self.colors["text"]

        return s
    
    # 手順解説
    def add_flow_slide(self, data: Dict[str, Any]):
        s = self._new_slide(data)
        self._add_slide_title(s, data["title"])

        steps = data.get("steps", [])
        body = data.get("bodyText", "")
        n = len(steps)
        direction = data.get("direction", "horizontal")

        slide_w, slide_h = self.prs.slide_width, self.prs.slide_height
        margin = Pt(60)
        spacing = Pt(50)

        if direction == "horizontal":
            # 横フロー
            box_w = (slide_w - margin*2 - spacing*(n-1)) / n
            box_h = Pt(120)
            top = slide_h/2 - box_h/2
            left = margin

            for i, text in enumerate(steps):
                # ラウンドボックス
                shape = s.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    left, top, box_w, box_h
                )
                shape.fill.solid()
                shape.fill.fore_color.rgb = self.colors["surface"]
                shape.line.color.rgb = self.colors["primary"]

                # 数字（左上に重ねる）
                num_box = s.shapes.add_textbox(left-Pt(20), top-Pt(40), Pt(40), Pt(40))
                tf_num = num_box.text_frame
                tf_num.text = str(i+1)
                p_num = tf_num.paragraphs[0]
                run_num = p_num.runs[0]
                run_num.font.size = Pt(55)
                run_num.font.bold = True
                run_num.font.color.rgb = self.colors["accent"]
                p_num.alignment = PP_ALIGN.LEFT

                # 本文
                tf = shape.text_frame
                tf.text = text
                p = tf.paragraphs[0]
                run = p.runs[0]
                run.font.size = Pt(20)
                run.font.name = "BIZ UDPゴシック"
                run.font.color.rgb = self.colors["text"]  
                p.alignment = PP_ALIGN.CENTER

                # 矢印
                if i < n-1:
                    arrow = s.shapes.add_shape(
                        MSO_SHAPE.RIGHT_ARROW,
                        left+box_w+5, top+box_h/3, spacing-10, Pt(40)
                    )
                    arrow.fill.solid()
                    arrow.fill.fore_color.rgb = self.colors["accent"]
                    arrow.line.fill.background()

                left += box_w + spacing

            # BodyText（下）
            if body:
                tbox = s.shapes.add_textbox(
                    Pt(60), top+box_h+spacing, slide_w-Pt(120), Pt(100)
                )
                tf = tbox.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                self._style_text(p, body, self.fonts["sizes"]["body"], color=self.colors["text"])

        else:
            # 縦フロー（左寄せ）
            flow_area_w = slide_w * 0.55   # 左2/3
            body_area_left = slide_w * 0.65
            box_w = flow_area_w * 0.9      # さらに少し狭く
            box_h = (slide_h - margin*2 - spacing*(n-1)) / n
            left = margin + Pt(40)
            top = margin + Pt(40)

            for i, text in enumerate(steps):
                # ラウンドボックス
                shape = s.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    left, top, box_w, box_h
                )
                shape.fill.solid()
                shape.fill.fore_color.rgb = self.colors["surface"]
                shape.line.color.rgb = self.colors["primary"]

                # 数字（ボックスの左外）
                num_box = s.shapes.add_textbox(left-Pt(60), top, Pt(50), box_h)
                tf_num = num_box.text_frame
                tf_num.text = str(i+1)
                p_num = tf_num.paragraphs[0]
                run_num = p_num.runs[0]
                run_num.font.size = Pt(36)
                run_num.font.bold = True
                run_num.font.color.rgb = self.colors["accent"]
                p_num.alignment = PP_ALIGN.CENTER

                # 本文
                tf = shape.text_frame
                tf.text = text
                p = tf.paragraphs[0]
                run = p.runs[0]
                run.font.size = Pt(20)
                run.font.name = "BIZ UDPゴシック"
                run.font.color.rgb = self.colors["text"] 
                p.alignment = PP_ALIGN.CENTER

                # 矢印（下向き）
                if i < n-1:
                    arrow = s.shapes.add_shape(
                        MSO_SHAPE.DOWN_ARROW,
                        left + box_w / 2 - Pt(20), top+box_h+5, Pt(40), spacing-10
                    )
                    arrow.fill.solid()
                    arrow.fill.fore_color.rgb = self.colors["accent"]
                    arrow.line.fill.background()

                top += box_h + spacing

            # BodyText（右1/3）
            if body:
                body_box = s.shapes.add_textbox(
                    body_area_left, margin + Pt(40), slide_w - body_area_left - Pt(40),
                    slide_h - margin*2
                )
                tf = body_box.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                self._style_text(p, body, self.fonts["sizes"]["body"], color=self.colors["text"])

    def add_highlight(self, data: Dict[str, Any]):
        """
        type: "highlight"
        title: スライドタイトル
        keyword: 強調するキーワードや公式
        description: 下に配置する解説文
        """
        s = self._new_slide(data)

        self._add_slide_title(s, data["title"])

        slide_w, slide_h = self.prs.slide_width, self.prs.slide_height
        keyword = data.get("keyword", "")
        description = data.get("description", "")

        # ---------------- キーワードボックス ----------------
        font_size = 44
        line_height = font_size * 1.4
        max_chars_per_line = 14   # 1行あたりの想定文字数（日本語ベース）
        n_lines = math.ceil(len(keyword) / max_chars_per_line)

        box_w = slide_w * 0.7      # 幅は広め固定
        box_h = Pt(line_height * n_lines + 60)  # 行数に応じて高さ調整
        left = (slide_w - box_w) / 2
        top = slide_h * 0.35

        shape = s.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, left, top, box_w, box_h
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = self.colors["surface"]
        shape.line.color.rgb = self.colors["accent"]

        tf = shape.text_frame
        tf.text = keyword
        p = tf.paragraphs[0]
        run = p.runs[0]
        run.font.size = Pt(font_size)
        run.font.name = "BIZ UDPゴシック"
        run.font.bold = True
        run.font.color.rgb = self.colors["primary"]
        p.alignment = PP_ALIGN.CENTER

        # ---------------- 解説文 ----------------
        if description:
            desc_top = top + box_h + Pt(30)
            tbox = s.shapes.add_textbox(
                Pt(60), desc_top, slide_w - Pt(120), Pt(120)
            )
            tf_desc = tbox.text_frame
            tf_desc.word_wrap = True
            tf_desc.text = description
            p2 = tf_desc.paragraphs[0]
            run2 = p2.runs[0]
            run2.font.size = Pt(20)
            run2.font.name = "BIZ UDPゴシック"
            run2.font.color.rgb = self.colors["text"]
            p2.alignment = PP_ALIGN.CENTER    

    def add_hero(self, data: dict):
        s = self._new_slide(data)

        # スライドサイズ
        slide_w, slide_h = self.prs.slide_width, self.prs.slide_height

        # 背景画像（あれば適用）
        bg_url = data.get("background-image")
        if bg_url:
            stream, _ = self._load_image(bg_url)
            if stream:
                s.shapes.add_picture(stream, 0, 0, width=slide_w, height=slide_h)
        else:
            # 背景色だけ
            fill = s.background.fill
            fill.solid()
            fill.fore_color.rgb = self.colors["background"]

        # --- 半透明オーバーレイ ---
        overlay = s.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, 0, 0, slide_w, slide_h
        )
        overlay.fill.solid()
        overlay.fill.fore_color.rgb = RGBColor(0, 0, 0)  # 黒
        overlay.fill.fore_color.transparency = 0.5              # 30%透過
        overlay.line.fill.background()                   # 枠線なし
        self._set_shape_transparency(overlay, 40000)  # 40%ほど透過

        # --- タイトル（中央配置） ---
        tbox = s.shapes.add_textbox(0, slide_h*0.35, slide_w, slide_h*0.2)
        tf = tbox.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tp = tf.paragraphs[0]
        self._style_text(
            tp,
            data.get("title", ""),
            self.fonts["sizes"]["title"],
            bold=True,
            color=RGBColor(255, 255, 255),
            align=PP_ALIGN.CENTER
        )

        # --- 仕切り線 ---
        line_top = slide_h * 0.51
        line_left = slide_w * 0.1
        line_width = slide_w * 0.8
        shape = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, line_left, line_top, line_width, Pt(1))
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
        shape.line.fill.background()

        # --- サブタイトル（タイトルのすぐ下） ---
        subtitle = data.get("subtitle")
        if subtitle:
            stbox = s.shapes.add_textbox(0, line_top + Pt(10), slide_w, slide_h*0.15)
            stf = stbox.text_frame
            stf.vertical_anchor = MSO_ANCHOR.TOP
            sp = stf.paragraphs[0]
            self._style_text(
                sp,
                subtitle,
                self.fonts["sizes"]["subhead"],
                color=RGBColor(230, 230, 230),
                align=PP_ALIGN.CENTER
            )

        return s
    
    def add_quote(self, data: dict):
        s = self._new_slide(data,False)
        slide_w, slide_h = self.prs.slide_width, self.prs.slide_height
        side_w = slide_w / 3

        # --- 左半分（画像 or primaryカラー） ---
        bg_url = data.get("image")
        if bg_url:
            stream, im = self._load_image(bg_url)
            if stream:
                iw, ih = im.size
                slide_h = self.prs.slide_height

                # 縦をスライドにフィット
                scale = slide_h / ih
                new_w = iw * scale
                new_h = ih * scale
                
                s.shapes.add_picture(stream, 0, 0, width=new_w, height=new_h)

            # 半透明オーバーレイ
            overlay = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, side_w, slide_h)
            overlay.fill.solid()
            overlay.fill.fore_color.rgb = self.colors["ghost"]
            self._set_shape_transparency(overlay, 40000)  # 40%透過
            overlay.line.fill.background()
            overlay.shadow.inherit = False
        else:
            # 画像がなければprimaryカラーで塗りつぶし
            rect = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, side_w, slide_h)
            rect.fill.solid()
            rect.fill.fore_color.rgb = self.colors["ghost"]
            rect.line.fill.background()
            rect.shadow.inherit = False

        # --- 右半分（背景はスライド標準色） ---
        rect_right = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, side_w, 0, side_w * 2, slide_h)
        rect_right.fill.solid()
        rect_right.fill.fore_color.rgb = self.colors["background"]
        rect_right.line.fill.background()
        rect_right.shadow.inherit = False

        # --- 引用符アイコン ---
        quote_box = s.shapes.add_textbox(side_w + Pt(40), Pt(40), side_w - Pt(80), Pt(80))
        qf = quote_box.text_frame
        qp = qf.paragraphs[0]
        qp.text = "“"   # フォント依存でシャープな形を狙う
        run = qp.runs[0]
        run.font.name = "Arial"   
        run.font.size = Pt(150)
        run.font.bold = True
        run.font.color.rgb = self.colors["ghost"]
        qp.alignment = PP_ALIGN.LEFT

        # --- 引用文 ---
        qbox = s.shapes.add_textbox(side_w + Pt(80), slide_h*0.25, side_w * 2 - Pt(160), slide_h*0.4)
        tf = qbox.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        self._style_text(
            p,
            data.get("quote", ""),
            self.fonts["sizes"]["sectionTitle"],
            bold=True,
            color=self.colors["text"],
            align=PP_ALIGN.LEFT
        )

        # --- 引用元 ---
        author = data.get("author", "")
        if author:
            abox = s.shapes.add_textbox(side_w + Pt(80), slide_h*0.75, side_w * 2 - Pt(80), Pt(80))
            atf = abox.text_frame
            ap = atf.paragraphs[0]
            self._style_text(
                ap,
                author,
                self.fonts["sizes"]["subhead"],
                color=self.colors["subtext"],
                align=PP_ALIGN.LEFT
            )

        return s

    # ---------------------- ビルド関数 ----------------------
def build_pptx_from_plan(plan: Dict[str, Any], out_path: str):
    sf = SlideFactory(plan)   # plan を渡すように変更
    for spec in plan.get("slides", []):
        t = spec.get("type")
        if t=="title"        : sf.add_title(spec)
        elif t=="section"    : sf.add_section(spec)
        elif t=="content"    : sf.add_content(spec)
        elif t=="compare"    : sf.add_compare(spec)
        elif t=="cards"      : sf.add_cards(spec)
        elif t=="progress"   : sf.add_progress(spec)
        elif t=="timeline"   : sf.add_timeline(spec)
        elif t=="image-auto" : sf.add_image_auto(spec)
        elif t=="qa-question": sf.add_qa_question(spec)
        elif t=="qa-answer"  : sf.add_qa_answer(spec)
        elif t=="table"      : sf.add_table_slide(spec)
        elif t=="flow"       : sf.add_flow_slide(spec)
        elif t=="highlight"  : sf.add_highlight(spec)
        elif t=="hero"       : sf.add_hero(spec)
        elif t=="quote"      : sf.add_quote(spec)
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
