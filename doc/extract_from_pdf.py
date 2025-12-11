#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
使用 PyMuPDF 從 PDF 文件中提取 PCB 術語
並建立完整的 JSON 格式術語資料庫
"""

import fitz  # PyMuPDF
import json
import re
from opencc import OpenCC
from collections import defaultdict

def extract_text_from_pdf(pdf_path):
    """使用 PyMuPDF 提取 PDF 文本"""
    print(f'正在讀取 PDF：{pdf_path}')

    doc = fitz.open(pdf_path)
    all_text = []

    print(f'PDF 總頁數：{len(doc)} 頁')
    print('開始提取文本...')

    for page_num in range(len(doc)):
        page = doc[page_num]
        text = page.get_text()
        all_text.append(text)

        # 顯示進度
        if (page_num + 1) % 5 == 0:
            print(f'  已處理：{page_num + 1}/{len(doc)} 頁')

    doc.close()
    print('PDF 文本提取完成！\n')

    return '\n'.join(all_text)

def clean_text(text):
    """清理文本，移除多餘的空白和特殊字符"""
    # 移除頁碼等
    text = re.sub(r'^\d+$', '', text, flags=re.MULTILINE)
    # 移除多餘空行
    text = re.sub(r'\n\s*\n', '\n', text)
    return text

def parse_terminology_advanced(text):
    """
    進階解析術語
    支持多種格式：
    1. English 中文
    2. English (Additional Info) 中文
    3. English 中文1，中文2
    """
    terminology_dict = {}
    duplicate_count = 0

    lines = text.split('\n')
    print(f'總行數：{len(lines)}')
    print('開始解析術語...\n')

    for line in lines:
        line = line.strip()
        if not line or len(line) < 3:
            continue

        # 跳過純中文或純英文的行
        if not re.search(r'[A-Za-z]', line) or not re.search(r'[\u4e00-\u9fff]', line):
            continue

        # 嘗試匹配多種格式
        patterns = [
            # 格式1: English 中文 (最常見)
            r'^([A-Za-z0-9\s\(\)\-,\.\/\[\]\'\"\&\+]+?)\s{2,}([^\x00-\x7F]+.*)$',
            # 格式2: English 中文 (至少一個空格)
            r'^([A-Za-z0-9\s\(\)\-,\.\/\[\]\'\"\&\+]+?)\s+([^\x00-\x7F]+.*)$',
        ]

        matched = False
        for pattern in patterns:
            match = re.match(pattern, line)
            if match:
                english = match.group(1).strip()
                chinese_simplified = match.group(2).strip()

                # 過濾掉太短的或無效的條目
                if len(english) < 2 or len(chinese_simplified) < 1:
                    continue

                # 檢查是否重複
                if english in terminology_dict:
                    duplicate_count += 1
                    # 如果中文翻譯更詳細，則更新
                    if len(chinese_simplified) > len(terminology_dict[english]['simplified']):
                        terminology_dict[english]['simplified'] = chinese_simplified
                else:
                    terminology_dict[english] = {
                        'english': english,
                        'simplified': chinese_simplified,
                        'traditional': ''  # 稍後轉換
                    }

                matched = True
                break

        if not matched and re.search(r'[A-Za-z]', line) and re.search(r'[\u4e00-\u9fff]', line):
            # 記錄無法解析的行（供調試用）
            pass

    print(f'解析完成！')
    print(f'  - 找到術語：{len(terminology_dict)} 個')
    print(f'  - 重複條目：{duplicate_count} 個')
    print()

    return terminology_dict

def convert_to_traditional(terminology_dict, cc):
    """將簡體中文轉換為繁體中文"""
    print('開始轉換簡體為繁體...')

    for english, trans in terminology_dict.items():
        simplified = trans['simplified']
        traditional = cc.convert(simplified)
        trans['traditional'] = traditional

    print('轉換完成！\n')
    return terminology_dict

def create_reverse_lookup(terminology_dict):
    """創建反向查詢字典（簡體 -> 繁體）"""
    print('建立反向查詢字典...')
    reverse_dict = {}

    for english, trans in terminology_dict.items():
        simp = trans['simplified']
        trad = trans['traditional']

        # 加入完整詞條
        reverse_dict[simp] = trad

        # 處理用逗號、頓號分隔的多個翻譯
        separators = ['，', '、', ',']
        for sep in separators:
            if sep in simp:
                simp_parts = simp.split(sep)
                trad_parts = trad.split(sep)

                for s, t in zip(simp_parts, trad_parts):
                    s = s.strip()
                    t = t.strip()
                    if s and t and len(s) > 1:  # 避免單字符
                        reverse_dict[s] = t

    print(f'建立了 {len(reverse_dict)} 個詞條\n')
    return reverse_dict

def save_json(data, filename):
    """儲存為 JSON 文件"""
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f'✓ 已儲存：{filename}')

def analyze_terminology(terminology_dict):
    """分析術語統計資訊"""
    print('=' * 80)
    print('術語統計分析')
    print('=' * 80)
    print()

    # 按首字母統計
    letter_count = defaultdict(int)
    for eng in terminology_dict.keys():
        first_char = eng[0].upper() if eng else '?'
        if first_char.isalpha():
            letter_count[first_char] += 1
        else:
            letter_count['#'] += 1

    print('按首字母分布：')
    print('-' * 80)
    for letter in sorted(letter_count.keys()):
        count = letter_count[letter]
        bar = '█' * (count // 3)
        print(f'  {letter}: {count:4d} {bar}')
    print()

    # 詞長統計
    lengths = [len(eng) for eng in terminology_dict.keys()]
    avg_length = sum(lengths) / len(lengths) if lengths else 0
    max_length = max(lengths) if lengths else 0
    min_length = min(lengths) if lengths else 0

    print('術語長度統計：')
    print('-' * 80)
    print(f'  平均長度：{avg_length:.1f} 個字符')
    print(f'  最長術語：{max_length} 個字符')
    print(f'  最短術語：{min_length} 個字符')
    print()

    # 找出最長的術語
    longest_terms = sorted(terminology_dict.items(), key=lambda x: len(x[0]), reverse=True)[:5]
    print('最長的 5 個術語：')
    print('-' * 80)
    for i, (eng, trans) in enumerate(longest_terms, 1):
        print(f'  {i}. {eng} ({len(eng)} 字符)')
        print(f'     {trans["traditional"]}')
    print()

def show_samples(terminology_dict, count=20):
    """顯示術語範例"""
    print('=' * 80)
    print(f'術語範例（隨機顯示 {count} 個）')
    print('=' * 80)
    print()

    import random
    items = list(terminology_dict.items())
    samples = random.sample(items, min(count, len(items)))

    for i, (eng, trans) in enumerate(samples, 1):
        print(f'{i:2d}. 英文：{eng}')
        print(f'    簡體：{trans["simplified"]}')
        print(f'    繁體：{trans["traditional"]}')
        print()

def main():
    print('=' * 80)
    print('PCB 術語 PDF 提取工具（使用 PyMuPDF）')
    print('=' * 80)
    print()

    # PDF 文件路徑
    pdf_path = '/Users/hhhsiao/Desktop/work/document_translation/PCB-Terminology-English-vs-Chinese.pdf'

    # 步驟 1: 提取 PDF 文本
    print('【步驟 1】從 PDF 提取文本')
    print('-' * 80)
    raw_text = extract_text_from_pdf(pdf_path)
    print(f'提取的文本總長度：{len(raw_text)} 個字符\n')

    # 步驟 2: 清理文本
    print('【步驟 2】清理文本')
    print('-' * 80)
    cleaned_text = clean_text(raw_text)
    print(f'清理後的文本長度：{len(cleaned_text)} 個字符\n')

    # 步驟 3: 解析術語
    print('【步驟 3】解析術語')
    print('-' * 80)
    terminology = parse_terminology_advanced(cleaned_text)

    # 步驟 4: 轉換為繁體
    print('【步驟 4】簡體轉繁體')
    print('-' * 80)
    cc = OpenCC('s2t')
    terminology = convert_to_traditional(terminology, cc)

    # 步驟 5: 建立反向查詢字典
    print('【步驟 5】建立反向查詢字典')
    print('-' * 80)
    reverse_lookup = create_reverse_lookup(terminology)

    # 步驟 6: 儲存 JSON 文件
    print('【步驟 6】儲存 JSON 文件')
    print('-' * 80)
    save_json(terminology, 'pcb_terms_from_pdf.json')
    save_json(reverse_lookup, 'simp_to_trad_from_pdf.json')
    print()

    # 步驟 7: 統計分析
    analyze_terminology(terminology)

    # 步驟 8: 顯示範例
    show_samples(terminology, 20)

    # 最終總結
    print('=' * 80)
    print('提取完成！')
    print('=' * 80)
    print()
    print('生成的文件：')
    print(f'  1. pcb_terms_from_pdf.json - 完整術語資料庫 ({len(terminology)} 個術語)')
    print(f'  2. simp_to_trad_from_pdf.json - 簡繁對照字典 ({len(reverse_lookup)} 個詞條)')
    print()
    print('=' * 80)

if __name__ == '__main__':
    main()
