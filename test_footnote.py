# -*- coding: utf-8 -*-
"""测试脚注和尾注功能"""
from converter import MarkdownToWordConverter
from parser import parse_inline, InlineType

markdown_text = """# 测试脚注和尾注

这是一段包含脚注的文字[^1]。

这是一段包含尾注的文字[^^1]。

还有更多脚注[^2]和尾注[^^2]。

[^1]: 这是第一个脚注的内容。
[^2]: 这是第二个脚注的内容。
[^^1]: 这是第一个尾注的内容（显示在文档末尾）。
[^^2]: 这是第二个尾注的内容（显示在文档末尾）。
"""

# 测试 parse_inline 是否能识别尾注引用
test_line = "包含脚注[^1]和尾注[^^1]"
print(f"测试行: {test_line}")
elements = parse_inline(test_line)
print(f"解析结果: {len(elements)} 个元素")
for i, elem in enumerate(elements):
    print(f"  [{i}] type={elem.type}, content='{elem.content}'")

print("\n" + "="*50 + "\n")

try:
    converter = MarkdownToWordConverter()
    doc = converter.convert_text(markdown_text)
    converter.save('test_endnote_output.docx')
    print("成功生成文档: test_endnote_output.docx")
    print(f"收集到的脚注: {converter.footnotes}")
    print(f"收集到的尾注: {converter.endnotes}")
    print(f"待注入脚注: {getattr(converter, '_pending_footnotes', [])}")
    print(f"待注入尾注: {getattr(converter, '_pending_endnotes', [])}")
except Exception as e:
    import traceback
    traceback.print_exc()
    print(f"错误: {e}")
