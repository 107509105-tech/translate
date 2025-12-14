"""Textbox and flowchart processing"""

from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Font size constants (in half-points)
TEXTBOX_FONT_SIZE_SMALL = 14  # 7pt
TEXTBOX_FONT_SIZE_NORMAL = 16  # 8pt
TEXTBOX_LENGTH_THRESHOLD = 7  # Character count threshold for choosing font size


class TextboxProcessor:
    """Handles textbox/flowchart translation and sizing"""

    def __init__(self, translation_client):
        """Initialize with translation client

        Args:
            translation_client: TranslationClient instance for API calls
        """
        self.translation_client = translation_client

    def process(self, doc):
        """Translate all textboxes in document

        Args:
            doc: Document object from docx
        """
        # 取得文件主體元素
        body = doc.element.body
        if body is None:
            return

        # 尋找文件中所有的文字方塊內容元素（w:txbxContent 是文字方塊內容的 XML 標籤）
        textboxes = body.findall('.//' + qn('w:txbxContent'))

        if not textboxes:
            return

        # 處理每個文字方塊
        for textbox in textboxes:
            # 步驟 1: 翻譯文字方塊中的所有文字
            # 尋找所有文字元素（w:t 是文字內容的 XML 標籤）
            text_elements = textbox.findall('.//' + qn('w:t'))
            translated_text = ""

            # 逐一翻譯每個文字元素
            for text_elem in text_elements:
                if text_elem.text and text_elem.text.strip():
                    original_text = text_elem.text
                    translated_text = self.translation_client.translate(original_text)
                    if translated_text:
                        # 將原文替換為翻譯後的文字
                        text_elem.text = translated_text

            # 步驟 2: 根據翻譯後文字的長度決定適當的字體大小
            # 如果文字長度超過閾值，使用較小字體（7pt），否則使用正常字體（8pt）
            if len(translated_text) > TEXTBOX_LENGTH_THRESHOLD:
                half_points = TEXTBOX_FONT_SIZE_SMALL  # 7pt
            else:
                half_points = TEXTBOX_FONT_SIZE_NORMAL  # 8pt

            # 步驟 3: 將字體大小統一套用到整個文字方塊
            # 尋找所有文字段落（w:r 是文字段落的 XML 標籤，用於包含格式設定）
            runs = textbox.findall('.//' + qn('w:r'))
            for r in runs:
                # 取得或建立文字段落屬性元素（w:rPr）
                rPr = r.find(qn('w:rPr'))
                if rPr is None:
                    rPr = OxmlElement('w:rPr')
                    r.append(rPr)

                # 取得或建立字體大小元素（w:sz）
                sz = rPr.find(qn('w:sz'))
                if sz is None:
                    sz = OxmlElement('w:sz')
                    rPr.append(sz)

                # 設定字體大小值（以半點為單位，因此 14 = 7pt，16 = 8pt）
                sz.set(qn('w:val'), str(half_points))
