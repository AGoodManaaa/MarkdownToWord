# -*- coding: utf-8 -*-
"""测试新功能"""

# 测试1: Markdown格式规范化
from utils import normalize_markdown
test1 = '# 标题\n内容\n## 二级标题\n更多内容'
result1 = normalize_markdown(test1)
print('测试1 规范化:', '通过' if '\n\n' in result1 else '失败')

# 测试2: 公式格式转换
from utils import convert_latex_delimiters
test2 = r'行内公式\(x^2\)和块级公式\[y=mx+b\]'
result2 = convert_latex_delimiters(test2)
has_inline = '$x^2$' in result2
has_block = '$$y=mx+b$$' in result2
print('测试2 公式转换:', '通过' if has_inline and has_block else '失败')

# 测试3: 表格转义竖线
from parser import _split_table_cells
test3 = '| a \\| b | c |'
result3 = _split_table_cells(test3)
print('测试3 转义竖线:', '通过' if 'a | b' in result3 else '失败', result3)

# 测试4: 删除线
from parser import parse_inline, InlineType
test4 = '这是~~删除线~~文本'
result4 = parse_inline(test4)
has_strike = any(e.type == InlineType.STRIKETHROUGH for e in result4)
print('测试4 删除线:', '通过' if has_strike else '失败')

print('\n所有功能测试完成!')
