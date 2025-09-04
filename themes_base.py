# themes_base.py

from abc import ABC, abstractmethod
from pptx.util import Pt

class SlideTheme(ABC):

    def delete_default_title(self, slide):
        #DefaultTitleの削除
        for shape in slide.shapes:
            if shape.name in ("TitleText", "TitleBar"):
                sp = shape._element
                sp.getparent().remove(sp)

    def render_image_auto(self, factory, data):
        """画像の枚数に応じて適切なメソッドを呼び分ける"""
        slide = factory._new_slide(data)
        images = data.get("images", [])
        count = len(images)
        if count == 0:
            return
        
        if count == 1:
            return self.render_image1(factory, slide, images, Pt(24))
        elif count == 2:
            return self.render_image2(factory, slide, images, Pt(22))
        elif count == 3:
            return self.render_image3(factory, slide, images, Pt(18))
        elif count == 4:
            return self.render_image4(factory, slide, images, Pt(18))
        else:
            raise ValueError(f"Unsupported image count: {count}")

    @abstractmethod
    def render_title(self, factory, data):
        pass

    @abstractmethod
    def render_section(self, factory, data):
        pass

    @abstractmethod
    def render_content(self, factory, data):
        pass

    @abstractmethod
    def render_cards(self, factory, data):
        pass

    @abstractmethod
    def render_compare(self, factory, data):
        pass

    @abstractmethod
    def render_progress(self, factory, data):
        pass

    @abstractmethod
    def render_timeline(self, factory, data):
        pass

    @abstractmethod
    def render_image1(self, factory, slide, images, font_size):
        pass

    @abstractmethod
    def render_image2(self, factory, slide, images, font_size):
        pass

    @abstractmethod
    def render_image3(self, factory, slide, images, font_size):
        pass
    
    @abstractmethod
    def render_image4(self, factory, slide, images, font_size):
        pass

    @abstractmethod
    def render_qa_question(self,factory, data):
        pass

    @abstractmethod
    def render_qa_answer(self, factory, data):
        pass

    @abstractmethod
    def render_table(self, factory, data):
        pass

    @abstractmethod
    def render_flow(self, factory, data):
        pass

    @abstractmethod
    def render_highlight(self, factory, data):
        pass

    @abstractmethod
    def render_quote(self, factory, data):
        pass

    @abstractmethod
    def render_hero(self, factory, data):
        pass

    @abstractmethod
    def render_features(self, factory, data):
        pass

    @abstractmethod
    def render_closing(self, factory, data):
        pass



