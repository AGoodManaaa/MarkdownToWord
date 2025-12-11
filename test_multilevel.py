# -*- coding: utf-8 -*-
"""测试多级有序列表"""

test_md = """
1. 提取圆心
2. 寻找邻居
   1. 对于成百上千个点
   2. 代码使用了KD-Tree
3. 确定P
   1. 对于每个光纤中心
   2. 这两个点之间的距离
   3. 注：在蜂窝结构中
"""

from parser import parse_markdown

blocks = parse_markdown(test_md)
for block in blocks:
    if block.type == 'list':
        print(f"列表类型: {'有序' if block.level == 1 else '无序'}")
        for item in block.content:
            print(f"  级别{item.get('level', 0)}: {item.get('text', item)}")

print("\n测试通过!")
