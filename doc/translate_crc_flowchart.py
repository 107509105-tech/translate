#!/usr/bin/env python3
"""
CRC 流程图文档翻译工具
将 Word 文档中的中文流程图翻译成英文，并添加到原文档后面
"""

from docx import Document
import copy
import sys
import os


# 完整的中英文翻译映射表
TRANSLATION_MAP = {
    # 基本符号
    'N': 'N',
    'Y': 'Y',

    # 流程控制
    '开始': 'Start',
    '结束': 'End',

    # CRC 操作
    'CRC寄存器初始化': 'Initialize CRC register',
    'CRC右移一位': 'Shift CRC right by 1 bit',
    '为：0xffff': 'To: 0xffff',
    'MSB位补0': 'Fill MSB with 0',

    # 数据处理
    '输入:一个8位二进制数据': 'Input: An 8-bit binary data',
    '存入寄存器': 'Store in register',
    '与CRC初始低八位': 'With CRC initial lower 8 bits',
    '进行异或运算': 'Perform XOR operation',
    '与多项式A001进行异或运算': 'XOR with polynomial A001',

    # 判断条件
    'LSB移出位是否为1?': 'Is LSB shifted out bit 1?',
    'LSB移出位是否为1？': 'Is LSB shifted out bit 1?',
    '是否右移8次?': 'Has it shifted 8 times?',
    '是否右移8次？': 'Has it shifted 8 times?',
    '是否进行下一位8位数据处理': 'Process next 8-bit data?',

    # 输出
    '输出：CRC码': 'Output: CRC code',
}


def translate_text(text):
    """
    翻译文本

    Args:
        text: 待翻译的中文文本

    Returns:
        翻译后的英文文本，如果没有对应翻译则返回原文
    """
    text = text.strip()
    return TRANSLATION_MAP.get(text, text)


def translate_element_text(element):
    """
    递归翻译 XML 元素中的所有文本内容

    Args:
        element: XML 元素对象
    """
    for child in element.iter():
        # 查找所有的文本元素 (w:t)
        if child.tag.endswith('t'):
            if child.text and child.text.strip():
                translated = translate_text(child.text)
                if translated != child.text:
                    print(f"  翻译: '{child.text}' -> '{translated}'")
                child.text = translated


def add_page_break(doc):
    """
    添加分页符

    Args:
        doc: Document 对象
    """
    doc.add_page_break()


def process_document(input_file, output_file=None):
    """
    处理文档：复制原始内容并添加翻译版本

    Args:
        input_file: 输入文件路径
        output_file: 输出文件路径，如果为 None 则覆盖原文件
    """
    # 检查输入文件是否存在
    if not os.path.exists(input_file):
        print(f"错误: 文件 '{input_file}' 不存在！")
        sys.exit(1)

    # 如果没有指定输出文件，则覆盖原文件
    if output_file is None:
        output_file = input_file

    print(f"正在处理文档: {input_file}")
    print("=" * 60)

    # 读取文档
    try:
        doc = Document(input_file)
    except Exception as e:
        print(f"错误: 无法读取文档 - {e}")
        sys.exit(1)

    print("\n步骤 1: 读取原始文档结构...")
    original_body = doc.element.body
    original_elements = list(original_body)
    print(f"  找到 {len(original_elements)} 个元素")

    print("\n步骤 2: 添加分页符...")
    add_page_break(doc)

    print("\n步骤 3: 复制并翻译流程图...")
    translated_count = 0

    for element in original_elements:
        # 深度复制元素
        new_element = copy.deepcopy(element)

        # 翻译新元素中的所有文本
        translate_element_text(new_element)

        # 添加到文档末尾
        doc.element.body.append(new_element)
        translated_count += 1

    print(f"  已复制并翻译 {translated_count} 个元素")

    print("\n步骤 4: 保存文档...")
    try:
        doc.save(output_file)
        print(f"  成功保存到: {output_file}")
    except Exception as e:
        print(f"错误: 无法保存文档 - {e}")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("完成！")
    print(f"\n文档结构:")
    print(f"  第 1 页: 原始中文流程图")
    print(f"  第 2 页: 英文翻译版本流程图")

    print(f"\n翻译对照表:")
    print("-" * 60)
    for cn, en in sorted(TRANSLATION_MAP.items()):
        if cn not in ['N', 'Y']:  # 跳过单字母
            print(f"  {cn:30s} -> {en}")


def main():
    """主函数"""
    print("CRC 流程图文档翻译工具")
    print("=" * 60)

    # 默认输入文件
    input_file = 'CRC校验流程图.docx'

    # 检查命令行参数
    if len(sys.argv) > 1:
        input_file = sys.argv[1]

    # 处理文档
    process_document(input_file)

    print("\n使用方法:")
    print(f"  python {sys.argv[0]} [输入文件路径]")
    print(f"  如果不指定输入文件，默认处理 'CRC校验流程图.docx'")


if __name__ == '__main__':
    main()
