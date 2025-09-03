# themes_simplenote.py
import themes_base
import slides_section
import slides_content
import slides_cards
import slides_compare
import slides_progress
import slides_timeline
import slides_image1
import slides_image2
import slides_image3
import slides_image4
import slides_qa_question
import slides_qa_answer
import slides_table
import slides_flow
import slides_highlight
import slides_quote
import slides_hero

from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor


class SimpleNoteTheme(themes_base.SlideTheme):


    def add_full_height_image(self, factory, slide):
        stream, im = factory._load_image("simplenote1.png")
        if stream is None:
            print("[WARN] 画像を読み込めません: simplenote1.png")
            return None

        # スライド全体の高さ
        slide_height = factory.prs.slide_height

        # 画像を追加
        slide.shapes.add_picture(stream, 0, 0, height=slide_height)
    
    def top_title(title_str):
        pass

    def render_title(self, factory, data):
        slide = factory._new_slide(data)
        
        # 横棒
        self.add_full_height_image(factory,slide)

        #subject
        subject = data.get("subject")
        if subject:
            sbox = slide.shapes.add_textbox(Pt(100), Pt(170), factory.prs.slide_width, Pt(20))
            sp = sbox.text_frame.paragraphs[0]
            factory._style_text(
                sp,
                subject,
                Pt(24),
                color=factory.colors["text"]
            )
            sbox.text_frame.word_wrap = True
            

        # タイトル
        head = data.get("title")
        if head:
            sbox = slide.shapes.add_textbox(Pt(100), Pt(200), factory.prs.slide_width, Pt(50))
            sp = sbox.text_frame.paragraphs[0]
            factory._style_text(
                sp,
                head,
                factory.fonts["sizes"]["contentTitle"],
                color=factory.colors["text"],
                bold=True
            )
            sbox.text_frame.word_wrap = True
        # タイトル下の横線
        line_top = Pt(260)
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Pt(100), line_top, Pt(300), Pt(2)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(0, 0, 0)  # 黒
        line.line.fill.background()  # 枠線なし
        line.shadow.inherit = False

        # タイトル下の横線2
        line_top = Pt(260)
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Pt(100), line_top + Pt(1), Pt(900), Pt(1)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(0, 0, 0)  # 黒
        line.line.fill.background()  # 枠線なし
        line.shadow.inherit = False

        # 講師名（固定）
        l_rect = factory.layout.get_rect("titleSlide.lecturer")
        lbox = slide.shapes.add_textbox(Pt(100), l_rect["top"], l_rect["width"], l_rect["height"])
        lp = lbox.text_frame.paragraphs[0]
        factory._style_text(
            lp,
            "講師名：〇〇　〇〇",
            factory.fonts["sizes"]["body"],
            color=factory.colors["text"]
        )
        # 日付
        d_rect = factory.layout.get_rect("titleSlide.date")
        dbox = slide.shapes.add_textbox(Pt(100), d_rect["top"], d_rect["width"], d_rect["height"])
        dp = dbox.text_frame.paragraphs[0]
        factory._style_text(
            dp,
            data.get("date", ""),
            factory.fonts["sizes"]["body"],
            color=factory.colors["subtext"]
        )





    def render_section(self, factory, data):
        return slides_section.render_section_default(factory, data)
    
    def render_content(self, factory, data):
        return slides_content.render_content_default(factory, data)
    
    def render_cards(self, factory, data):
        return slides_cards.render_cards_default(factory, data)
    
    def render_compare(self, factory, data):
        return slides_compare.render_compare_default(factory, data)
    
    def render_progress(self, factory, data):
        return slides_progress.render_progress_default(factory, data)
    
    def render_timeline(self, factory, data):
        return slides_timeline.render_timeline_default(factory, data)
    
    def render_image1(self, factory, slide, images, font_size):
        return slides_image1.render_image1_default(factory, slide, images, font_size)

    def render_image2(self, factory, slide, images, font_size):
        return slides_image2.render_image2_default(factory, slide, images, font_size)

    def render_image3(self, factory, slide, images, font_size):
        return slides_image3.render_image3_default(factory, slide, images, font_size)

    def render_image4(self, factory, slide, images, font_size):
        return slides_image4.render_image4_default(factory, slide, images, font_size)

    def render_qa_question(self, factory, data):
        return slides_qa_question.render_qa_question_defaults(factory, data)
    
    def render_qa_answer(self, factory, data):
        return slides_qa_answer.render_qa_answer_default(factory, data)

    def render_table(self, factory, data):
        return slides_table.render_table_default(factory, data)
    
    def render_flow(self, factory, data):
        return slides_flow.render_flow_default(factory, data)

    def render_highlight(self, factory, data):
        return slides_highlight.render_highlight_default(factory, data)

    def render_quote(self, factory, data):
        return slides_quote.render_quote_default(factory, data)

    def render_hero(self, factory, data):
        return slides_hero.render_hero_default(factory, data)
    


