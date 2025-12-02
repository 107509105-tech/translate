import fitz  # PyMuPDF
import google.generativeai as genai
import time
import re
import os

# 配置 Google AI API
# 请设置环境变量 GOOGLE_API_KEY 或在这里直接填入
api_key = os.getenv('GOOGLE_API_KEY', 'YOUR_API_KEY_HERE')
genai.configure(api_key=api_key)

# 初始化 Gemma 模型
# 使用 Gemini 1.5 Flash 作为快速翻译模型（也可以换成其他模型）
model = genai.GenerativeModel('gemini-1.5-flash')

# 翻译缓存
translation_cache = {}

def has_chinese(text):
    """检查文本是否包含中文字符"""
    return bool(re.search('[\u4e00-\u9fff]', text))

def translate_text(text):
    """使用 Gemma/Gemini 翻译文本，使用缓存避免重复翻译"""
    if not text or not text.strip():
        return text

    # 如果不包含中文，直接返回
    if not has_chinese(text):
        return text

    # 标准化文本用于缓存
    text_key = text.strip()

    # 检查缓存
    if text_key in translation_cache:
        return translation_cache[text_key]

    # 使用 Gemini/Gemma 翻译
    try:
        # 构建翻译提示
        prompt = f"""Translate the following Traditional Chinese text to English.
Only output the translated text without any explanations or additional comments.

Text to translate: {text_key}

Translation:"""

        response = model.generate_content(prompt)
        translated = response.text.strip()

        # 去除可能的引号或额外格式
        translated = translated.strip('"\'')

        translation_cache[text_key] = translated
        print(f"  Translated: '{text_key[:40]}...' -> '{translated[:40]}...'")
        time.sleep(0.2)  # 避免请求过快
        return translated
    except Exception as e:
        print(f"  Translation error for '{text_key[:40]}...': {e}")
        return text

def convert_color(color_int):
    """将整数颜色值转换为RGB元组"""
    if isinstance(color_int, int):
        r = ((color_int >> 16) & 0xFF) / 255.0
        g = ((color_int >> 8) & 0xFF) / 255.0
        b = (color_int & 0xFF) / 255.0
        return (r, g, b)
    return (0, 0, 0)

def create_translated_pdf(input_pdf, output_pdf):
    """创建翻译后的PDF"""
    print(f"\nOpening PDF: {input_pdf}")
    doc = fitz.open(input_pdf)

    total_translated = 0
    total_spans = 0

    for page_num in range(len(doc)):
        print(f"\n{'='*80}")
        print(f"Processing Page {page_num + 1}...")
        print('='*80)

        page = doc[page_num]

        # 获取所有文本块
        text_dict = page.get_text("dict")
        replacements = []

        for block in text_dict["blocks"]:
            if block["type"] == 0:  # 文本块
                for line in block["lines"]:
                    for span in line["spans"]:
                        total_spans += 1
                        original = span["text"]

                        if has_chinese(original):
                            translated = translate_text(original)

                            if translated != original:
                                total_translated += 1
                                replacements.append({
                                    "rect": fitz.Rect(span["bbox"]),
                                    "original": original,
                                    "translated": translated,
                                    "fontsize": span["size"],
                                    "color": span["color"],
                                    "flags": span["flags"]
                                })

        print(f"\nApplying {len(replacements)} translations to page {page_num + 1}...")

        # 应用替换
        for repl in replacements:
            # 用白色矩形覆盖原文
            page.draw_rect(repl["rect"], color=(1, 1, 1), fill=(1, 1, 1))

            # 转换颜色
            color = convert_color(repl["color"])

            # 插入翻译文本
            # 根据文本长度动态调整字体大小
            rect_width = repl["rect"].width
            rect_height = repl["rect"].height
            text_length = len(repl["translated"])

            # 估算合适的字体大小
            if text_length > 100:
                fontsize = repl["fontsize"] * 0.5
            elif text_length > 50:
                fontsize = repl["fontsize"] * 0.65
            else:
                fontsize = repl["fontsize"] * 0.75

            # 确保字体不会太小
            fontsize = max(fontsize, 6)

            try:
                rc = page.insert_textbox(
                    repl["rect"],
                    repl["translated"],
                    fontsize=fontsize,
                    fontname="helv",
                    color=color,
                    align=fitz.TEXT_ALIGN_LEFT,
                )

                # 如果文本太长无法适应，继续缩小字体
                attempts = 0
                while rc < 0 and attempts < 3:
                    fontsize *= 0.8
                    rc = page.insert_textbox(
                        repl["rect"],
                        repl["translated"],
                        fontsize=fontsize,
                        fontname="helv",
                        color=color,
                        align=fitz.TEXT_ALIGN_LEFT,
                    )
                    attempts += 1

                if rc < 0:
                    print(f"  Warning: Text too long to fit: '{repl['translated'][:40]}...'")

            except Exception as e:
                print(f"  Error inserting text: {e}")

    # 保存
    print(f"\n{'='*80}")
    print("Saving translated PDF...")
    doc.save(output_pdf, garbage=4, deflate=True)
    doc.close()

    print(f"{'='*80}")
    print(f"\n✓ Translation completed successfully!")
    print(f"  Total text spans: {total_spans}")
    print(f"  Translated spans: {total_translated}")
    print(f"  Translation coverage: {total_translated/total_spans*100:.1f}%")
    print(f"  Output saved to: {output_pdf}")
    print(f"{'='*80}\n")

if __name__ == "__main__":
    input_file = "113_chartN.pdf"
    output_file = "113_chartN_english_gemma.pdf"

    print("\n" + "="*80)
    print("PDF TRANSLATION TOOL - With Gemini/Gemma LLM")
    print("="*80)
    print(f"Input file:  {input_file}")
    print(f"Output file: {output_file}")
    print("="*80)

    # 执行翻译
    create_translated_pdf(input_file, output_file)

    print("\n✓ All done! Please check the output file.")
