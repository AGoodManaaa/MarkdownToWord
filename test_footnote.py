# -*- coding: utf-8 -*-
"""测试脚注功能 - 带调试输出"""
from converter import MarkdownToWordConverter
from parser import parse_inline, InlineType

markdown_text = """# 测试脚注

这是一段包含脚注的文字[^1]。

还有另一个脚注[^2]。

[^1]: 这是第一个脚注的内容。
[^2]: 这是第二个脚注的内容。
"""

# 测试 parse_inline 是否能识别脚注引用
test_line = "这是一段包含脚注的文字[^1]。"
print(f"测试行: {test_line}")
elements = parse_inline(test_line)
print(f"解析结果: {len(elements)} 个元素")
for i, elem in enumerate(elements):
    print(f"  [{i}] type={elem.type}, content='{elem.content}'")

print("\n" + "="*50 + "\n")

try:
    converter = MarkdownToWordConverter()
    doc = converter.convert_text(markdown_text)
    converter.save('test_footnote_output.docx')
    print("成功生成文档: test_footnote_output.docx")
    print(f"收集到的脚注: {converter.footnotes}")
    print(f"待注入脚注: {getattr(converter, '_pending_footnotes', [])}")
except Exception as e:
    import traceback
    traceback.print_exc()
    print(f"错误: {e}")
