import fitz  # PyMuPDF
import json
from typing import Dict, List, Tuple
import re

# 翻译字典 - 关键术语的翻译映射
TRANSLATION_DICT = {
    # 标题
    "碩士班研究生修業流程圖": "Master's Program Graduate Study Flowchart",

    # 表头
    "作業流程": "Process",
    "時間": "Timeline",
    "說明": "Description",

    # 流程步骤
    "先修課程修習／抵免": "Prerequisite Courses / Credit Exemption",
    "選定指導教授": "Advisor Selection",
    "論文修習": "Thesis Study",
    "學術倫理": "Academic Ethics",
    "提交論文研究計畫": "Thesis Research Proposal Submission",
    "論文研究計畫報告": "Thesis Research Proposal Report",
    "學位考試（口試）申請": "Degree Examination (Oral Defense) Application",
    "論文口試": "Thesis Oral Defense",
    "離校申請": "Graduation Checkout",

    # 时间描述
    "入學前(依教務處公告時間辦理)": "Before Enrollment (According to Academic Affairs Office Schedule)",
    "入學後第二學期結束前完成": "Complete Before End of Second Semester After Enrollment",
    "入學期間": "During Enrollment",
    "依進度提供": "According to Progress",
    "計畫書審查通過後": "After Proposal Review Approval",
    "3 個月": "3 Months",
    "審查通過後2個月": "2 Months After Review Approval",
    "報告審查通過經 2 個月，並且於口試舉行前一週或當學期學位考試申請截止日前(各學期學位考試申請截止日依教務處規定)；口試日期訂於申請截止日之後者，仍需在截止日期前提出申請。": "2 months after report approval, and one week before the oral examination or before the degree examination application deadline of the current semester (deadline determined by Academic Affairs Office); if the oral examination is scheduled after the application deadline, application must still be submitted before the deadline.",
    "申請學位考試當學期結束日前舉行完畢。（每年度1/30或7/30前）": "Must be completed before the end of the semester in which the degree examination is applied. (Before January 30 or July 30 annually)",
    "未申請延後畢業學生應於學位考試當學期結束日後一個月內（上學期 3/1、下學期 9/1）": "Students not applying for delayed graduation should complete within one month after the end of the semester of degree examination (Spring semester: March 1, Fall semester: September 1)",

    # 常用词汇
    "先修課程": "Prerequisite Courses",
    "非招生考試且非資訊相關學生": "Non-entrance exam and non-information related students",
    "資料結構": "Data Structures",
    "演算法": "Algorithms",
    "作業系統": "Operating Systems",
    "計算機組織與結構": "Computer Organization and Architecture",
    "學分抵免": "Credit Exemption",
    "指導教授": "Advisor",
    "論文題目": "Thesis Title",
    "論文計畫": "Thesis Proposal",
    "專題研討": "Seminar",
    "畢業學分": "Graduation Credits",
    "學術倫理": "Academic Ethics",
    "學位考試": "Degree Examination",
    "論文計畫書": "Thesis Proposal",
    "進度報告": "Progress Report",
    "口試": "Oral Defense",
    "口試委員": "Examination Committee Members",
    "成績報告表": "Grade Report Form",
    "審定書": "Approval Certificate",
    "論文電子檔": "Thesis Electronic File",
    "離校程序": "Graduation Checkout Procedure",
    "學位證書": "Degree Certificate",
    "系辦": "Department Office",
    "教務處": "Academic Affairs Office",
    "註冊組": "Registration Office",
    "圖書館": "Library",
    "校務系統": "University Information System",
    "修業規定": "Curriculum Regulations",
    "系務會議": "Department Affairs Meeting",
    "論文研究報告": "Thesis Research Report",
    "論文抄襲比對系統": "Thesis Plagiarism Detection System",
    "畢業初審": "Preliminary Graduation Review",
    "複審": "Secondary Review",
    "邀請函": "Invitation Letter",
    "系戳": "Department Seal",
    "精裝本": "Hardcover",
    "平裝本": "Paperback",
    "學生證": "Student ID",
    "秘書室": "Secretary Office",
    "生僑組": "Student Affairs Office",
    "國合處": "International Cooperation Office",
}

def translate_text(text: str) -> str:
    """
    翻译中文文本到英文
    优先使用字典中的翻译，然后进行智能翻译
    """
    # 首先查找完全匹配
    if text in TRANSLATION_DICT:
        return TRANSLATION_DICT[text]

    # 处理包含多个部分的文本
    translated = text

    # 按照长度降序排序，优先匹配更长的短语
    sorted_keys = sorted(TRANSLATION_DICT.keys(), key=len, reverse=True)

    for chinese, english in [(k, TRANSLATION_DICT[k]) for k in sorted_keys]:
        if chinese in translated:
            translated = translated.replace(chinese, english)

    # 处理剩余的常见模式
    translations = {
        "學分": "credits",
        "個月": "months",
        "學期": "semester",
        "上學期": "spring semester",
        "下學期": "fall semester",
        "或": "or",
        "及": "and",
        "應": "should",
        "須": "must",
        "需": "need",
        "可": "can",
        "至": "to",
        "於": "at/in",
        "依": "according to",
        "經": "after",
        "並": "and",
        "以": "with",
        "完成": "complete",
        "申請": "application/apply",
        "送": "submit",
        "繳交": "submit",
        "辦理": "process",
        "進行": "proceed",
        "詳見": "see",
        "規定": "regulations",
        "說明": "instructions",
        "網址": "URL",
        "官網": "official website",
        "下載專區": "download area",
        "申請表格": "application forms",
        "碩士班": "master's program",
        "學生": "students",
        "同意": "approval/agree",
        "簽名": "signature/sign",
        "簽章": "sign and seal",
        "紙本": "paper copy",
        "印出": "print",
        "雙面列印": "double-sided printing",
        "附件": "attachment",
        "資料": "materials/data",
        "相關": "related",
        "問題": "questions",
        "可洽": "contact",
    }

    for cn, en in translations.items():
        if cn in translated and cn not in TRANSLATION_DICT:
            # 只翻译还没有被翻译的部分
            translated = translated.replace(cn, f" {en} ")

    return translated.strip()

def extract_text_with_format(page) -> List[Dict]:
    """
    提取页面中的文本及其格式信息
    """
    blocks = []
    text_dict = page.get_text("dict")

    for block in text_dict["blocks"]:
        if block["type"] == 0:  # 文本块
            for line in block["lines"]:
                for span in line["spans"]:
                    blocks.append({
                        "text": span["text"],
                        "bbox": span["bbox"],  # (x0, y0, x1, y1)
                        "font": span["font"],
                        "size": span["size"],
                        "color": span["color"],
                        "flags": span["flags"],  # 字体样式标志
                    })
        elif block["type"] == 1:  # 图像块
            blocks.append({
                "type": "image",
                "bbox": block["bbox"],
                "image": block.get("image"),
            })

    return blocks

def create_translated_pdf(input_pdf: str, output_pdf: str):
    """
    创建翻译后的PDF，保持原有格式
    """
    # 打开原始PDF
    doc = fitz.open(input_pdf)

    # 创建新的PDF
    new_doc = fitz.open()

    for page_num in range(len(doc)):
        page = doc[page_num]

        # 创建新页面，使用相同的尺寸
        new_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)

        # 首先复制所有图形元素（背景、线条、形状等）
        new_page.show_pdf_page(new_page.rect, doc, page_num, clip=page.rect)

        # 提取文本块
        blocks = extract_text_with_format(page)

        # 在新页面上绘制翻译后的文本
        for block in blocks:
            if block.get("type") == "image":
                continue  # 跳过图像，因为已经通过show_pdf_page复制了

            text = block["text"]
            translated = translate_text(text)

            if not translated.strip():
                continue

            bbox = block["bbox"]
            x0, y0, x1, y1 = bbox

            # 计算文本框的尺寸
            width = x1 - x0
            height = y1 - y0

            # 获取字体信息
            fontsize = block["size"]

            # 根据flags确定字体样式
            flags = block["flags"]
            is_bold = flags & 2**4
            is_italic = flags & 2**1

            # 选择字体
            if is_bold and is_italic:
                fontname = "helv-oblique"  # Helvetica Bold Italic
            elif is_bold:
                fontname = "helv"  # Helvetica Bold
            elif is_italic:
                fontname = "helv-oblique"  # Helvetica Italic
            else:
                fontname = "helv"  # Helvetica Regular

            # 调整字体大小以适应框
            # 英文通常比中文占用更多水平空间，可能需要缩小
            adjusted_fontsize = fontsize * 0.8  # 初始调整

            try:
                # 插入文本
                # 首先在原位置画白色矩形覆盖原文
                new_page.draw_rect(fitz.Rect(x0, y0, x1, y1), color=(1, 1, 1), fill=(1, 1, 1))

                # 插入翻译后的文本
                new_page.insert_textbox(
                    fitz.Rect(x0, y0, x1, y1),
                    translated,
                    fontsize=adjusted_fontsize,
                    fontname=fontname,
                    color=block.get("color", 0),  # 使用原始颜色
                    align=fitz.TEXT_ALIGN_LEFT,
                )
            except Exception as e:
                print(f"Error inserting text '{translated}': {e}")
                continue

    # 保存新PDF
    new_doc.save(output_pdf)
    new_doc.close()
    doc.close()

    print(f"Translation completed! Output saved to: {output_pdf}")

def extract_and_save_text(input_pdf: str, output_json: str):
    """
    提取PDF中的所有文本并保存为JSON，方便检查和手动调整翻译
    """
    doc = fitz.open(input_pdf)
    all_text = []

    for page_num in range(len(doc)):
        page = doc[page_num]
        blocks = extract_text_with_format(page)

        for block in blocks:
            if block.get("type") == "image":
                continue

            text = block["text"]
            translated = translate_text(text)

            all_text.append({
                "original": text,
                "translated": translated,
                "page": page_num + 1,
                "bbox": block["bbox"],
            })

    doc.close()

    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(all_text, f, ensure_ascii=False, indent=2)

    print(f"Text extracted and saved to: {output_json}")

if __name__ == "__main__":
    input_file = "113_chartN.pdf"
    output_file = "113_chartN_english.pdf"
    json_file = "translation_preview.json"

    # 首先提取文本并保存为JSON，方便检查
    print("Extracting text...")
    extract_and_save_text(input_file, json_file)

    # 创建翻译后的PDF
    print("Creating translated PDF...")
    create_translated_pdf(input_file, output_file)

    print("\nDone! Please check the following files:")
    print(f"1. {json_file} - Preview of all translations")
    print(f"2. {output_file} - Translated PDF")
