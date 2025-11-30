import fitz  # PyMuPDF
import re

# 完整的翻译映射表
TRANSLATIONS = {
    # 主标题
    "碩士班研究生修業流程圖": "Master's Program Graduate Study Flowchart",

    # 表格标题
    "作業流程": "Process Flow",
    "時間": "Timeline",
    "說明": "Description",

    # 流程步骤
    "先修課程修習／抵免": "Prerequisite Courses\nCompletion/Exemption",
    "先修課程修": "Prerequisite Courses",
    "習／抵免": "Completion/Exemption",

    "選定指導教授": "Select Academic\nAdvisor",
    "選定指導": "Select Academic",
    "教授": "Advisor",

    "論文修習": "Thesis Study",
    "學術倫理": "Academic Ethics",

    "提交論文研究計畫": "Submit Thesis\nResearch Proposal",
    "提交論文": "Submit Thesis",
    "研究計畫": "Research Proposal",

    "論文研究計畫報告": "Thesis Research\nProposal Report",
    "論文研究": "Thesis Research",
    "計畫報告": "Proposal Report",

    "審查通過後3個月": "3 months after review",
    "審查通過後2個月": "2 months after review",

    "學位考試（口試）申請": "Degree Examination\n(Oral Defense)\nApplication",
    "學位考試（口試）": "Degree Examination\n(Oral Defense)",
    "申請": "Application",

    "論文口試": "Thesis Oral\nDefense",

    "離校申請": "Graduation\nCheckout",

    # 时间说明
    "入學前(依教務處公告時間辦理)": "Before Enrollment\n(According to Academic Affairs Office Schedule)",
    "入學前(依教務處公告": "Before Enrollment",
    "時間辦理)": "(According to Schedule)",

    "入學後第二學期結束前完成": "Complete Before the End of\nSecond Semester After Enrollment",
    "入學後第二學期結束": "Before End of Second Semester",
    "前完成": "After Enrollment",

    "入學期間": "During Enrollment",
    "依進度提供": "According to Progress",

    "計畫書審查通過後": "After Proposal Review",
    "3 個月": "3 Months",

    "報告審查通過經 2 個月，並且於口試舉行前一週或當學期學位考試申請截止日前(各學期學位考試申請截止日依教務處規定)；口試日期訂於申請截止日之後者，仍需在截止日期前提出申請。": "2 months after report approval, and one week before oral examination or before degree examination application deadline of current semester (deadline determined by Academic Affairs Office); if oral examination is scheduled after application deadline, application must still be submitted before deadline.",

    "報告審查通過經 2 個": "2 months after",
    "月，並且於口試舉行": "report approval, and",
    "前一週或當學期學位": "one week before oral",
    "考試申請截止日前(各": "examination or before",
    "學期學位考試申請截止": "degree examination",
    "日依教務處規定)；口試": "application deadline",
    "日期訂於申請截止日": "(per Academic Affairs",
    "之後者，仍需在截止": "Office); if scheduled",
    "日期前提出申請。": "after, apply before deadline.",

    "申請學位考試當學期結束日前舉行完畢。（每年度1/30或7/30前）": "Must complete before end of semester of degree examination application.\n(Before Jan 30 or Jul 30 annually)",
    "申請學位考試當學期": "Must complete before end",
    "結束日前舉行完畢。": "of examination semester.",
    "（每年度1/30或7/30前）": "(Before Jan 30 or Jul 30)",

    "未申請延後畢業學生應於學位考試當學期結束日後一個月內（上學期 3/1、下學期 9/1）": "Students not applying for delayed graduation should complete within one month after end of degree examination semester\n(Spring: Mar 1, Fall: Sep 1)",
    "未申請延後畢業學生": "Students not applying for",
    "應於學位考試當學期": "delayed graduation should",
    "結束日後一個月內（上": "complete within one month",
    "學期 3/1、下學期 9/1）": "(Spring: 3/1, Fall: 9/1)",

    # 详细说明内容
    "1.先修課程：非招生考試且非資訊相關學生應自【資料結構、演算法、作業系統、計算機組織與結構】四科中選修二科通過─大學階段已修過者，抵免資料送系辦審查。": "1. Prerequisite Courses: Non-entrance exam and non-CS students must select and pass 2 of 4 courses [Data Structures, Algorithms, Operating Systems, Computer Organization & Architecture]. Those who completed these in undergraduate studies should submit exemption materials to department office for review.",

    "※先修抵免申請表：資科系官網/下載專區/申請表格-碩士班/先修科目抵免申請表。": "※Prerequisite Exemption Form: Department Website/Downloads/Forms-Master's Program/Prerequisite Course Exemption Form.",

    "2.學分抵免：本系畢業學分抵免上限為9 學分；依本校抵免辦法進行。": "2. Credit Exemption: Maximum exemption for graduation is 9 credits; processed per university exemption regulations.",

    "※學分抵免申請表：教務處註冊組/碩博士班新生學分抵免申請表": "※Credit Exemption Form: Academic Affairs Registration/Credit Exemption Form for New Master's & Doctoral Students",

    "1.選定指導教授，並至 iNCCU/校務系統 Web 版入口/校務資訊系統/學生資訊系統/學術服務/研究生申報論文題目，辦理申報並請指導教授簽名。": "1. Select advisor and go to iNCCU/University System Web Portal/Academic Info System/Student Info System/Academic Services/Graduate Thesis Title Declaration to process declaration and request advisor's signature.",

    "2.變更指導教授時，應取得新指導教授，以及原指導教授或系務會議同意。": "2. When changing advisors, must obtain approval from new advisor and original advisor or department affairs meeting.",

    "3.變更指導教授後，應重新申報論文計畫，該計畫應經新指導教授及原指導教授或系務會議同意。": "3. After changing advisors, must re-declare thesis proposal, which requires approval from new advisor and original advisor or department affairs meeting.",

    "1.依課務組選課作業流程辦理。": "1. Process according to Course Affairs Section course selection procedures.",

    "2.須修習專題研討（一）~（四），至少 2 學分。": "2. Must complete Seminar (I)~(IV), at least 2 credits.",

    "3.畢業學分須達 26 學分-修習外系課程之規定詳見修業規定。": "3. Must achieve 26 graduation credits - regulations for external courses detailed in curriculum regulations.",

    "4.需於碩班 Seminar 完成論文研究報告。": "4. Must complete thesis research report in Master's Program Seminar.",

    "學生須完成學術倫理始得申請學位考試。抵免問題可洽註冊組。學術研究倫理課修習說明可至系網查詢。": "Students must complete academic ethics before applying for degree examination. Credit exemption questions: contact Registration Office. Academic research ethics course info available on department website.",

    "※臺灣學術倫理教育資源中心網址：https://ethics.moe.edu.tw。": "※Taiwan Academic Ethics Education Resource Center: https://ethics.moe.edu.tw",

    "提交簡式論文計畫書至系辦；應登記 Seminer 時間，附件雙面列印。": "Submit simplified thesis proposal to department office; register Seminar time, attachments double-sided printed.",

    "※論文計畫書：資科系官網/下載專區/申請表格-碩士班/論文計畫書申請表。": "※Thesis Proposal: Department Website/Downloads/Forms-Master's Program/Thesis Proposal Application Form.",

    "提交進度報告(Progress Report)至系辦；應提供 Seminer 時間。附件雙面列印。": "Submit Progress Report to department office; provide Seminar time. Attachments double-sided printed.",

    "※進度報告：資科系官網/下載專區/申請表格-碩士班/碩士論文進度申請表。": "※Progress Report: Department Website/Downloads/Forms-Master's Program/Master's Thesis Progress Application Form.",

    "1.線上申請「學位考試」─至 iNCCU/校務系統 Web 版入口/校務資訊系統/學生資訊系統/學術服務/學位考試申請系統，進行個人化檢核後，需紙本印出送指導教授簽名。": "1. Apply for \"Degree Examination\" online - go to iNCCU/University System Web Portal/Academic Info System/Student Info System/Academic Services/Degree Exam Application System. After personalized verification, print paper copy for advisor's signature.",

    "2.論文需校內論文抄襲比對系統比對─將比對結果紙本印出，送請指導教授簽名同意，併同「學位考試申請書」送系所完成畢業初審並經系所主管簽章，申請書送教務處複審同意後，始得進行學位考試。": "2. Thesis must be checked through campus plagiarism detection system - print comparison results, submit to advisor for signature/approval, along with \"Degree Examination Application Form\" to complete preliminary graduation review by department and obtain department head's signature/seal. After Academic Affairs Office secondary review approval, degree examination may proceed.",

    "1.口試前：準備邀請函(需送系辦蓋系戳)、各委員一份成績報告表、學位考試成績報告單、中/英文審定書等資料。": "1. Before oral defense: Prepare invitation letters (need department seal), one grade report form per committee member, degree examination grade report, Chinese/English approval certificates, etc.",

    "2.口試後：學位考試成績報告單經口試委員評分並簽名，繳交至系辦。": "2. After oral defense: Degree examination grade report graded and signed by examination committee members, submit to department office.",

    "※學位考試相關資料：資科系官網/下載專區/學位考試。": "※Degree Examination Materials: Department Website/Downloads/Degree Examination.",

    "1. 上傳論文電子檔 (https://www.lib.nccu.edu.tw/p/404-1000-288.php?Lang=zh-tw)，詳見圖書館之規定。": "1. Upload thesis electronic file (https://www.lib.nccu.edu.tw/p/404-1000-288.php?Lang=zh-tw), see library regulations.",

    "2. 請至 iNCCU 列印畢業離校程序單，依離校程序單所示，完成系所、圖書館、秘書室、生僑組、國合處等相關單位之檢核。": "2. Go to iNCCU to print graduation checkout form. Per checkout form, complete verification by department, library, secretary office, student affairs, international cooperation, etc.",

    "3. 碩士論文精裝本或平裝本需繳交 2 冊存於圖書館，本系不留存。": "3. Submit 2 copies of master's thesis (hardcover or paperback) to library. Department does not keep copies.",

    "4. 完成離校程序後，即可攜帶學生證至註冊組領取學位證書。": "4. After completing checkout procedure, bring student ID to Registration Office to receive degree certificate.",
}

def normalize_text(text):
    """标准化文本用于匹配"""
    return re.sub(r'\s+', ' ', text.strip())

def translate_text(text):
    """翻译文本"""
    normalized = normalize_text(text)

    # 精确匹配
    if normalized in TRANSLATIONS:
        return TRANSLATIONS[normalized]

    # 尝试模糊匹配
    for key, value in TRANSLATIONS.items():
        if normalize_text(key) == normalized:
            return value

    # 返回原文（未找到翻译）
    return text

def create_translated_pdf(input_pdf, output_pdf):
    """创建翻译后的PDF"""
    doc = fitz.open(input_pdf)

    for page_num in range(len(doc)):
        page = doc[page_num]

        # 获取所有文本块
        text_dict = page.get_text("dict")
        replacements = []

        for block in text_dict["blocks"]:
            if block["type"] == 0:  # 文本块
                for line in block["lines"]:
                    for span in line["spans"]:
                        original = span["text"]
                        translated = translate_text(original)

                        if translated != original:
                            replacements.append({
                                "rect": fitz.Rect(span["bbox"]),
                                "original": original,
                                "translated": translated,
                                "fontsize": span["size"],
                                "color": span["color"],
                                "flags": span["flags"]
                            })

        # 应用替换
        for repl in replacements:
            # 用白色矩形覆盖原文
            page.draw_rect(repl["rect"], color=(1, 1, 1), fill=(1, 1, 1))

            # 转换颜色
            color_int = repl["color"]
            if isinstance(color_int, int):
                r = ((color_int >> 16) & 0xFF) / 255.0
                g = ((color_int >> 8) & 0xFF) / 255.0
                b = (color_int & 0xFF) / 255.0
                color = (r, g, b)
            else:
                color = (0, 0, 0)

            # 插入翻译文本
            fontsize = repl["fontsize"] * 0.75  # 缩小字体以适应英文

            try:
                rc = page.insert_textbox(
                    repl["rect"],
                    repl["translated"],
                    fontsize=fontsize,
                    fontname="helv",
                    color=color,
                    align=fitz.TEXT_ALIGN_LEFT,
                )

                # 如果文本太长，继续缩小字体
                if rc < 0:
                    page.insert_textbox(
                        repl["rect"],
                        repl["translated"],
                        fontsize=fontsize * 0.6,
                        fontname="helv",
                        color=color,
                        align=fitz.TEXT_ALIGN_LEFT,
                    )
            except Exception as e:
                print(f"Warning: Could not insert '{repl['translated'][:30]}...': {e}")

    # 保存
    doc.save(output_pdf, garbage=4, deflate=True)
    doc.close()
    print(f"✓ Translation completed successfully!")
    print(f"✓ Output saved to: {output_pdf}")

def print_translation_stats(input_pdf):
    """打印翻译统计"""
    doc = fitz.open(input_pdf)

    total_spans = 0
    translated_spans = 0
    untranslated = set()

    for page_num in range(len(doc)):
        page = doc[page_num]
        text_dict = page.get_text("dict")

        for block in text_dict["blocks"]:
            if block["type"] == 0:
                for line in block["lines"]:
                    for span in line["spans"]:
                        text = span["text"].strip()
                        if text:
                            total_spans += 1
                            translated = translate_text(text)
                            if translated != text:
                                translated_spans += 1
                            else:
                                # 只记录中文未翻译的
                                if any('\u4e00' <= c <= '\u9fff' for c in text):
                                    untranslated.add(text)

    doc.close()

    print("\n" + "="*80)
    print("TRANSLATION STATISTICS")
    print("="*80)
    print(f"Total text spans: {total_spans}")
    print(f"Translated spans: {translated_spans}")
    print(f"Translation coverage: {translated_spans/total_spans*100:.1f}%")

    if untranslated:
        print(f"\nUntranslated Chinese text ({len(untranslated)} unique items):")
        for text in sorted(untranslated)[:20]:  # 只显示前20个
            print(f"  - {text}")
        if len(untranslated) > 20:
            print(f"  ... and {len(untranslated)-20} more")
    print("="*80 + "\n")

if __name__ == "__main__":
    input_file = "113_chartN.pdf"
    output_file = "113_chartN_english.pdf"

    print("PDF Translation Tool - PyMuPDF")
    print("="*80)
    print(f"Input file:  {input_file}")
    print(f"Output file: {output_file}")
    print()

    # 打印翻译统计
    print_translation_stats(input_file)

    # 执行翻译
    print("Creating translated PDF...")
    create_translated_pdf(input_file, output_file)

    print("\n" + "="*80)
    print("✓ All done! Please check the output file.")
    print("="*80)
