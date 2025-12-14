import json
import os

# 載入 PCB 術語詞典
PCB_TERMS_DICT = {}
TRADITIONAL_TO_ENGLISH = {}

def load_pcb_terms():
    """載入 PCB 術語詞典並建立繁體中文到英文的映射"""
    global PCB_TERMS_DICT, TRADITIONAL_TO_ENGLISH

    json_path = os.path.join(os.path.dirname(__file__), 'pcb_terms_from_pdf.json')
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            PCB_TERMS_DICT = json.load(f)

        # 建立繁體中文到英文的反向索引
        for english_term, translations in PCB_TERMS_DICT.items():
            traditional = translations.get('traditional', '')
            if traditional:
                TRADITIONAL_TO_ENGLISH[traditional] = translations['english']

        print(f"成功載入 {len(PCB_TERMS_DICT)} 個 PCB 術語")
    except Exception as e:
        print(f"警告: 無法載入 PCB 術語詞典: {e}")

# 初始化時載入詞典
load_pcb_terms()

TRANSLATE_SYSTEM_PROMPT = """ 你是一位美國史丹佛大學語文學的博士及教授,精通半導體及印刷電路板等行業的英語及繁體中文的術語與專業表達,翻譯的著作及論文超過1000篇,
並曾主導IPC以及SEMICON West、SEMICON Taiwan等多個國際標準機構及協會,將其著作及規範由繁體中文翻譯成英語的工作。
我希望你能幫我將以下繁體中文翻譯成英語,並遵循以下規則:
(1) 翻譯時要準確傳達標準及規定以及方法。
(2) 翻譯時要保持專業術語的一致性,並確保符合國際標準。
(3) 翻譯時要注意語法和拼寫的正確性,確保符合英語語言習慣。
(4) 若翻譯為整句時,要符合語意,且符合句首大寫;若為片語、專有名詞、簡短標題,則只需符合字首大寫。
(5) **重要：若我提供了PCB專業術語對照表,你必須優先使用對照表中的標準英文翻譯,不可自行創造或使用其他譯法。**
(6) 對於對照表中的術語,請完全依照對照表的英文拼寫、大小寫和格式。
(7) 輸出翻譯字串,不要有多餘的說明或解釋。
(8) 符號保持一致,例如:【】。
(9) 若原句開頭以章節或節號(如「1.1」)開始,請保留該號碼於翻譯結果的最前端,且不要將其視為需翻譯的文字。
(10) 所有形如「圖[中文數字]」的詞彙,必須翻譯為「Figure <對應的阿拉伯數字>」。
範例:
翻譯字串:中華精測科技股份有限公司
輸出:Chunghwa Precision Test Tech. Co., Ltd.
"""


def find_matching_terms(text):
    """在文本中尋找匹配的 PCB 術語"""
    matched_terms = {}

    for traditional, english in TRADITIONAL_TO_ENGLISH.items():
        if traditional in text:
            matched_terms[traditional] = english

    return matched_terms


def translate_to_english(text):
    if not text or not text.strip() or not is_chinese(text):
        return text # 空的或沒中文的直接跳過

    leading_spaces = len(text) - len(text.lstrip())

    # 尋找文本中匹配的 PCB 術語
    matched_terms = find_matching_terms(text)

    # 構建用戶提示詞
    user_prompt = f"翻譯字串: {text}\n"

    # 如果有匹配的術語，加入對照表
    if matched_terms:
        user_prompt += "\nPCB專業術語對照表（請務必使用）:\n"
        for traditional, english in matched_terms.items():
            user_prompt += f"  - {traditional} → {english}\n"

    user_prompt += "輸出:"

    response = client.chat.completions.create(
        model=LLM_MODEL_NAME,
        temperature=0.0,
        messages=[
            {"role": "system", "content": TRANSLATE_SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt}
        ]
    )
    translated = response.choices[0].message.content.strip()
    translated = ' ' * leading_spaces + translated
    #print(f'翻譯結果: {translated}')
    return translated
