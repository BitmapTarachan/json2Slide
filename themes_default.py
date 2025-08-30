# themes_default.py
from themes_base import SlideTheme
import slides_title
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


class DefaultTheme(SlideTheme):

    def render_title(self, factory, data):
        return slides_title.render_title_default(factory, data)

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
    


