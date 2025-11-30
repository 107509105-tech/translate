import fitz  # PyMuPDF
import json

# 完整的翻译字典
FULL_TRANSLATIONS = {
    # 标题
    "碩士班研究生修業流程圖": "Master's Program Graduate Study Flowchart",

    # 表头
    "作業流程": "Process",
    "時間": "Timeline",
    "說明": "Description",

    # 流程步骤标题
    "先修課程修習／抵免": "Prerequisite Courses\nCompletion/Exemption",
    "選定指導教授": "Select Academic\nAdvisor",
    "論文修習": "Thesis Study",
    "學術倫理": "Academic Ethics",
    "提交論文研究計畫": "Submit Thesis\nResearch Proposal",
    "論文研究計畫報告": "Thesis Research\nProposal Report",
    "學位考試（口試）申請": "Degree Examination\n(Oral Defense)\nApplication",
    "論文口試": "Thesis Oral\nDefense",
    "離校申請": "Graduation\nCheckout",

    # 时间描述
    "入學前(依教務處公告時間辦理)": "Before Enrollment (According to Academic Affairs Office Schedule)",
    "入學後第二學期結束前完成": "Complete Before the End of the Second Semester After Enrollment",
    "入學期間": "During Enrollment",
    "依進度提供": "According to Progress",
    "計畫書審查通過後": "After Proposal Review",
    "3 個月": "3 Months",
    "審查通過後2個月": "2 Months After Review",
    "報告審查通過經 2 個月，並且於口試舉行前一週或當學期學位考試申請截止日前(各學期學位考試申請截止日依教務處規定)；口試日期訂於申請截止日之後者，仍需在截止日期前提出申請。": "2 months after report approval, and one week before the oral examination or before the degree examination application deadline of the current semester (deadline determined by Academic Affairs Office); if the oral examination is scheduled after the application deadline, application must still be submitted before the deadline.",
    "申請學位考試當學期結束日前舉行完畢。（每年度1/30或7/30前）": "Must be completed before the end of the semester in which the degree examination is applied. (Before January 30 or July 30 annually)",
    "未申請延後畢業學生應於學位考試當學期結束日後一個月內（上學期 3/1、下學期 9/1）": "Students not applying for delayed graduation should complete within one month after the end of the semester of degree examination (Spring semester: March 1, Fall semester: September 1)",

    # 说明内容 - 第一部分
    "1.先修課程：非招生考試且非資訊相關學生應自【資料結構、演算法、作業系統、計算機組織與結構】四科中選修二科通過─大學階段已修過者，抵免資料送系辦審查。": "1. Prerequisite Courses: Non-entrance exam and non-information related students should select and pass two courses from [Data Structures, Algorithms, Operating Systems, Computer Organization and Architecture]. Students who have completed these courses during undergraduate studies should submit exemption materials to the department office for review.",
    "※先修抵免申請表：資科系官網/下載專區/申請表格-碩士班/先修科目抵免申請表。": "※Prerequisite Exemption Application Form: Department Website/Download Area/Application Forms-Master's Program/Prerequisite Course Exemption Application Form.",
    "2.學分抵免：本系畢業學分抵免上限為9 學分；依本校抵免辦法進行。": "2. Credit Exemption: The maximum credit exemption for graduation from this department is 9 credits; processed according to the university's exemption regulations.",
    "※學分抵免申請表：教務處註冊組/碩博士班新生學分抵免申請表": "※Credit Exemption Application Form: Academic Affairs Office Registration Section/Credit Exemption Application Form for New Master's and Doctoral Students",

    # 说明内容 - 第二部分
    "1.選定指導教授，並至 iNCCU/校務系統 Web 版入口/校務資訊系統/學生資訊系統/學術服務/研究生申報論文題目，辦理申報並請指導教授簽名。": "1. Select an academic advisor, and go to iNCCU/University Information System Web Portal/Academic Information System/Student Information System/Academic Services/Graduate Thesis Title Declaration to process the declaration and request advisor's signature.",
    "2.變更指導教授時，應取得新指導教授，以及原指導教授或系務會議同意。": "2. When changing advisors, approval must be obtained from the new advisor, as well as the original advisor or the department affairs meeting.",
    "3.變更指導教授後，應重新申報論文計畫，該計畫應經新指導教授及原指導教授或系務會議同意。": "3. After changing advisors, the thesis proposal should be re-declared, which requires approval from both the new advisor and the original advisor or the department affairs meeting.",

    # 说明内容 - 第三部分
    "1.依課務組選課作業流程辦理。": "1. Process according to the course selection procedures of the Course Affairs Section.",
    "2.須修習專題研討（一）~（四），至少 2 學分。": "2. Must complete Seminar (I)~(IV), at least 2 credits.",
    "3.畢業學分須達 26 學分-修習外系課程之規定詳見修業規定。": "3. Graduation credits must reach 26 credits - regulations for taking courses outside the department can be found in the curriculum regulations.",
    "4.需於碩班 Seminar 完成論文研究報告。": "4. Must complete thesis research report in the Master's Program Seminar.",

    # 说明内容 - 第四部分
    "學生須完成學術倫理始得申請學位考試。抵免問題可洽註冊組。學術研究倫理課修習說明可至系網查詢。": "Students must complete academic ethics before applying for degree examination. Credit exemption questions can be directed to the Registration Office. Academic research ethics course information can be found on the department website.",
    "※臺灣學術倫理教育資源中心網址：https://ethics.moe.edu.tw。": "※Taiwan Academic Ethics Education Resource Center URL: https://ethics.moe.edu.tw.",

    # 说明内容 - 第五部分
    "提交簡式論文計畫書至系辦；應登記 Seminer 時間，附件雙面列印。": "Submit simplified thesis proposal to the department office; should register Seminar time, attachments printed double-sided.",
    "※論文計畫書：資科系官網/下載專區/申請表格-碩士班/論文計畫書申請表。": "※Thesis Proposal: Department Website/Download Area/Application Forms-Master's Program/Thesis Proposal Application Form.",

    # 说明内容 - 第六部分
    "提交進度報告(Progress Report)至系辦；應提供 Seminer 時間。附件雙面列印。": "Submit Progress Report to the department office; should provide Seminar time. Attachments printed double-sided.",
    "※進度報告：資科系官網/下載專區/申請表格-碩士班/碩士論文進度申請表。": "※Progress Report: Department Website/Download Area/Application Forms-Master's Program/Master's Thesis Progress Application Form.",

    # 说明内容 - 第七部分 (分段)
    "1.線上申請「學位考試」─至 iNCCU/校務系統 Web 版入口/校務資訊系統/學生資訊系統/學術服務/學位考試申請系統，進行個人化檢核後，需紙本印出送指導教授簽名。": "1. Apply for \"Degree Examination\" online - go to iNCCU/University Information System Web Portal/Academic Information System/Student Information System/Academic Services/Degree Examination Application System. After personalized verification, print out the paper copy for advisor's signature.",
    "2.論文需校內論文抄襲比對系統比對─將比對結果紙本印出，送請指導教授簽名同意，併同「學位考試申請書」送系所完成畢業初審並經系所主管簽章，申請書送教務處複審同意後，始得進行學位考試。": "2. Thesis must be checked through the campus plagiarism detection system - print out the comparison results, submit to advisor for signature and approval, along with the \"Degree Examination Application Form\" to complete the preliminary graduation review by the department and obtain the department head's signature and seal. After the application is approved by the Academic Affairs Office for secondary review, the degree examination may proceed.",

    # 说明内容 - 第八部分 (分段)
    "1.口試前：準備邀請函(需送系辦蓋系戳)、各委員一份成績報告表、學位考試成績報告單、中/英文審定書等資料。": "1. Before oral defense: Prepare invitation letters (need department office seal), one grade report form for each committee member, degree examination grade report, Chinese/English approval certificates, and other materials.",
    "2.口試後：學位考試成績報告單經口試委員評分並簽名，繳交至系辦。": "2. After oral defense: Degree examination grade report signed and graded by oral examination committee members, submit to department office.",
    "※學位考試相關資料：資科系官網/下載專區/學位考試。": "※Degree Examination Related Materials: Department Website/Download Area/Degree Examination.",

    # 说明内容 - 第九部分 (分段)
    "1. 上傳論文電子檔 (https://www.lib.nccu.edu.tw/p/404-1000-288.php?Lang=zh-tw)，詳見圖書館之規定。": "1. Upload thesis electronic file (https://www.lib.nccu.edu.tw/p/404-1000-288.php?Lang=zh-tw), see library regulations for details.",
    "2. 請至 iNCCU 列印畢業離校程序單，依離校程序單所示，完成系所、圖書館、秘書室、生僑組、國合處等相關單位之檢核。": "2. Go to iNCCU to print the graduation checkout procedure form. According to the checkout procedure form, complete the verification of relevant units such as the department, library, secretary office, student affairs office, and international cooperation office.",
    "3. 碩士論文精裝本或平裝本需繳交 2 冊存於圖書館，本系不留存。": "3. Two copies of the master's thesis (hardcover or paperback) must be submitted to the library. The department does not keep copies.",
    "4. 完成離校程序後，即可攜帶學生證至註冊組領取學位證書。": "4. After completing the checkout procedure, you can take your student ID to the Registration Office to receive the degree certificate.",
}

def convert_color(color_int):
    """
    将整数颜色值转换为RGB元组 (0-1范围)
    """
    if isinstance(color_int, int):
        # 将整数转换为RGB
        r = ((color_int >> 16) & 0xFF) / 255.0
        g = ((color_int >> 8) & 0xFF) / 255.0
        b = (color_int & 0xFF) / 255.0
        return (r, g, b)
    return (0, 0, 0)  # 默认黑色

def translate_text(text: str) -> str:
    """翻译文本"""
    # 去除首尾空格进行匹配
    text_stripped = text.strip()

    # 完全匹配
    if text_stripped in FULL_TRANSLATIONS:
        return FULL_TRANSLATIONS[text_stripped]

    # 如果没有找到翻译，返回原文
    return text

def create_translated_pdf_simple(input_pdf: str, output_pdf: str):
    """
    创建翻译后的PDF - 使用更简单的方法
    复制原PDF并覆盖文本
    """
    doc = fitz.open(input_pdf)

    for page_num in range(len(doc)):
        page = doc[page_num]

        # 获取文本块
        text_dict = page.get_text("dict")
        blocks_to_redact = []
        text_insertions = []

        for block in text_dict["blocks"]:
            if block["type"] == 0:  # 文本块
                for line in block["lines"]:
                    for span in line["spans"]:
                        text = span["text"]
                        translated = translate_text(text)

                        if translated and translated != text:
                            bbox = span["bbox"]

                            # 记录需要覆盖的区域
                            blocks_to_redact.append(fitz.Rect(bbox))

                            # 准备插入的文本
                            fontsize = span["size"]
                            color = convert_color(span["color"])

                            text_insertions.append({
                                "rect": fitz.Rect(bbox),
                                "text": translated,
                                "fontsize": fontsize * 0.85,  # 稍微缩小以适应英文
                                "color": color,
                            })

        # 用白色覆盖原文本
        for rect in blocks_to_redact:
            page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1))

        # 插入翻译后的文本
        for item in text_insertions:
            try:
                rc = page.insert_textbox(
                    item["rect"],
                    item["text"],
                    fontsize=item["fontsize"],
                    fontname="helv",
                    color=item["color"],
                    align=fitz.TEXT_ALIGN_LEFT,
                )
                if rc < 0:
                    # 如果文本太长，尝试更小的字体
                    page.insert_textbox(
                        item["rect"],
                        item["text"],
                        fontsize=item["fontsize"] * 0.7,
                        fontname="helv",
                        color=item["color"],
                        align=fitz.TEXT_ALIGN_LEFT,
                    )
            except Exception as e:
                print(f"Error inserting text '{item['text'][:30]}...': {e}")

    # 保存
    doc.save(output_pdf, garbage=4, deflate=True)
    doc.close()
    print(f"✓ Translation completed! Output saved to: {output_pdf}")

def show_translation_preview(input_pdf: str):
    """显示将要翻译的内容预览"""
    doc = fitz.open(input_pdf)

    print("\n" + "="*80)
    print("TRANSLATION PREVIEW")
    print("="*80)

    for page_num in range(len(doc)):
        page = doc[page_num]
        text_dict = page.get_text("dict")

        print(f"\n--- Page {page_num + 1} ---")

        for block in text_dict["blocks"]:
            if block["type"] == 0:
                for line in block["lines"]:
                    for span in line["spans"]:
                        text = span["text"].strip()
                        if text:
                            translated = translate_text(text)
                            if translated != text:
                                print(f"  CN: {text}")
                                print(f"  EN: {translated}")
                                print()

    doc.close()
    print("="*80)

if __name__ == "__main__":
    input_file = "113_chartN.pdf"
    output_file = "113_chartN_english.pdf"

    print("Starting PDF translation process...")
    print(f"Input: {input_file}")
    print(f"Output: {output_file}\n")

    # 显示翻译预览
    show_translation_preview(input_file)

    # 创建翻译后的PDF
    print("\nCreating translated PDF...")
    create_translated_pdf_simple(input_file, output_file)

    print(f"\n✓ Done! Translated PDF saved as: {output_file}")
