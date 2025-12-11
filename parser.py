# -*- coding: utf-8 -*-
"""
Markdown 解析器 - 共用模块
预览和Word导出使用同一套解析逻辑
"""

import re
from typing import List, Tuple, Any
from dataclasses import dataclass
from enum import Enum


class InlineType(Enum):
    """行内元素类型"""
    TEXT = "text"
    BOLD = "bold"
    ITALIC = "italic"
    BOLD_ITALIC = "bold_italic"
    CODE = "code"
    MATH = "math"
    LINK = "link"
    IMAGE = "image"
    STRIKETHROUGH = "strikethrough"
    SUPERSCRIPT = "superscript"  # 上标
    SUBSCRIPT = "subscript"      # 下标
    LINEBREAK = "linebreak"      # 换行 <br>


@dataclass
class InlineElement:
    """行内元素"""
    type: InlineType
    content: str
    url: str = None  # 用于链接和图片


@dataclass
class BlockElement:
    """块级元素"""
    type: str  # heading, paragraph, code_block, table, quote, list, math_block, hr, image
    content: Any  # 内容，可以是字符串、列表或其他结构
    level: int = 0  # 用于标题级别
    language: str = ""  # 用于代码块


def parse_inline(text: str) -> List[InlineElement]:
    """
    解析行内元素，返回元素列表
    这是预览和Word导出共用的核心解析逻辑
    使用简单可靠的顺序匹配方法
    """
    elements = []
    
    # 简化的正则模式（按优先级排序，使用非贪婪匹配）
    # 合并成一个大正则，每个模式用括号分组
    # 注意：斜体匹配需要确保不会匹配乘法表达式中的星号
    pattern = (
        r'(!\[[^\]]*\]\([^\)]+\))'           # 1: 图片
        r'|(\[[^\]]+\]\([^\)]+\))'            # 2: 链接
        r'|(`[^`]+`)'                          # 3: 行内代码
        r'|(\$[^$]+\$)'                        # 4: 行内公式
        r'|(<br\s*/?>)'                         # 5: 换行标签 <br> 或 <br/>
        r'|(<sup>[^<]+</sup>)'                 # 6: 上标 HTML
        r'|(<sub>[^<]+</sub>)'                 # 7: 下标 HTML
        r'|(\*\*\*.+?\*\*\*)'                  # 8: 粗斜体 ***
        r'|(___.+?___)'                        # 9: 粗斜体 ___
        r'|(\*\*.+?\*\*)'                      # 10: 粗体 **
        r'|(__.+?__)'                          # 11: 粗体 __
        r'|(?<!\*)\*(?!\*)([^\*\s][^\*]*[^\*\s]|[^\*\s])\*(?!\*)'  # 12: 斜体 *text*
        r'|(?<!_)_(?!_)([^_\s][^_]*[^_\s]|[^_\s])_(?!_)'    # 13: 斜体 _text_
        r'|(~~.+?~~)'                          # 14: 删除线
    )
    
    last_end = 0
    for match in re.finditer(pattern, text):
        # 添加匹配前的普通文本
        if match.start() > last_end:
            plain_text = text[last_end:match.start()]
            if plain_text:
                elements.append(InlineElement(InlineType.TEXT, plain_text))
        
        full_match = match.group(0)
        
        # 根据匹配内容判断类型
        if full_match.startswith('!['):
            # 图片 ![alt](url)
            m = re.match(r'!\[([^\]]*)\]\(([^\)]+)\)', full_match)
            if m:
                elements.append(InlineElement(InlineType.IMAGE, m.group(1), m.group(2)))
        
        elif full_match.startswith('['):
            # 链接 [text](url)
            m = re.match(r'\[([^\]]+)\]\(([^\)]+)\)', full_match)
            if m:
                elements.append(InlineElement(InlineType.LINK, m.group(1), m.group(2)))
        
        elif full_match.startswith('`'):
            # 行内代码 `code`
            content = full_match[1:-1]
            elements.append(InlineElement(InlineType.CODE, content))
        
        elif full_match.startswith('$'):
            # 行内公式 $formula$
            content = full_match[1:-1]
            elements.append(InlineElement(InlineType.MATH, content))
        
        elif full_match.startswith('<br'):
            # 换行标签 <br> 或 <br/>
            elements.append(InlineElement(InlineType.LINEBREAK, '\n'))
        
        elif full_match.startswith('<sup>'):
            # 上标 <sup>text</sup>
            content = full_match[5:-6]  # 去掉 <sup> 和 </sup>
            elements.append(InlineElement(InlineType.SUPERSCRIPT, content))
        
        elif full_match.startswith('<sub>'):
            # 下标 <sub>text</sub>
            content = full_match[5:-6]  # 去掉 <sub> 和 </sub>
            elements.append(InlineElement(InlineType.SUBSCRIPT, content))
        
        elif full_match.startswith('***') or full_match.startswith('___'):
            # 粗斜体
            content = full_match[3:-3]
            elements.append(InlineElement(InlineType.BOLD_ITALIC, content))
        
        elif full_match.startswith('**'):
            # 粗体 **text**
            content = full_match[2:-2]
            elements.append(InlineElement(InlineType.BOLD, content))
        
        elif full_match.startswith('__'):
            # 粗体 __text__
            content = full_match[2:-2]
            elements.append(InlineElement(InlineType.BOLD, content))
        
        elif full_match.startswith('~~'):
            # 删除线
            content = full_match[2:-2]
            elements.append(InlineElement(InlineType.STRIKETHROUGH, content))
        
        elif full_match.startswith('*'):
            # 斜体 *text* - 需要提取内部内容
            # 由于使用了新的正则，匹配结果可能在不同的组中
            inner = re.match(r'^\*([^\*]+)\*$', full_match)
            if inner:
                content = inner.group(1)
            else:
                content = full_match[1:-1] if len(full_match) > 2 else full_match
            elements.append(InlineElement(InlineType.ITALIC, content))
        
        elif full_match.startswith('_'):
            # 斜体 _text_
            inner = re.match(r'^_([^_]+)_$', full_match)
            if inner:
                content = inner.group(1)
            else:
                content = full_match[1:-1] if len(full_match) > 2 else full_match
            elements.append(InlineElement(InlineType.ITALIC, content))
        
        last_end = match.end()
    
    # 添加剩余文本
    if last_end < len(text):
        remaining = text[last_end:]
        if remaining:
            elements.append(InlineElement(InlineType.TEXT, remaining))
    
    # 如果没有匹配到任何元素，返回整个文本
    if not elements:
        elements.append(InlineElement(InlineType.TEXT, text))
    
    return elements


def parse_markdown(text: str) -> List[BlockElement]:
    """
    解析Markdown文本，返回块级元素列表
    """
    blocks = []
    lines = text.split('\n')
    i = 0
    
    while i < len(lines):
        line = lines[i]
        
        # 空行
        if not line.strip():
            i += 1
            continue
        
        # 代码块 ```
        if line.strip().startswith('```'):
            lang = line.strip()[3:].strip()
            code_lines = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i])
                i += 1
            i += 1  # 跳过结束的 ```
            blocks.append(BlockElement('code_block', '\n'.join(code_lines), language=lang))
            continue
        
        # 块级公式 $$
        if line.strip().startswith('$$'):
            # 使用正则匹配第一个完整公式，处理同一行多个公式的情况
            single_match = re.match(r'^\$\$(.+?)\$\$', line.strip())
            if single_match:
                formula = single_match.group(1).strip()
                blocks.append(BlockElement('math_block', formula))
                # 检查剩余部分是否还有公式
                remaining = line.strip()[single_match.end():].strip()
                if remaining.startswith('$$'):
                    lines[i] = remaining  # 修改当前行供后续处理
                else:
                    i += 1
                continue
            
            # 多行公式（以单独 $$ 开头）
            formula_lines = []
            i += 1
            while i < len(lines) and not lines[i].strip().endswith('$$'):
                formula_lines.append(lines[i])
                i += 1
            if i < len(lines):
                last = lines[i].strip()
                if last != '$$':
                    formula_lines.append(last[:-2])
            i += 1
            blocks.append(BlockElement('math_block', '\n'.join(formula_lines)))
            continue
        
        # 标题 #
        if line.startswith('#'):
            level = len(re.match(r'^#+', line).group())
            content = line[level:].strip()
            blocks.append(BlockElement('heading', content, level=min(level, 4)))
            i += 1
            continue
        
        # 表格（对分隔行进行 strip 处理，确保匹配）
        if '|' in line and i + 1 < len(lines) and re.match(r'^[\s\|\:\-]+$', lines[i + 1].strip()):
            table_lines = []
            while i < len(lines) and '|' in lines[i]:
                table_lines.append(lines[i])
                i += 1
            blocks.append(BlockElement('table', '\n'.join(table_lines)))
            continue
        
        # 引用 >
        if line.startswith('>'):
            quote_lines = []
            while i < len(lines) and lines[i].startswith('>'):
                quote_lines.append(lines[i][1:].strip())
                i += 1
            blocks.append(BlockElement('quote', '\n'.join(quote_lines)))
            continue
        
        # 无序列表（包括任务列表）
        if re.match(r'^[\s]*[\*\-\+]\s', line):
            items = []
            while i < len(lines) and re.match(r'^[\s]*[\*\-\+]\s', lines[i]):
                item_text = re.sub(r'^[\s]*[\*\-\+]\s*', '', lines[i])
                # 检查是否是任务列表
                task_match = re.match(r'^\[([ xX])\]\s*(.*)', item_text)
                if task_match:
                    checked = task_match.group(1).lower() == 'x'
                    text = task_match.group(2)
                    items.append({'type': 'task', 'checked': checked, 'text': text})
                else:
                    items.append({'type': 'item', 'text': item_text})
                i += 1
            blocks.append(BlockElement('list', items, level=0))
            continue
        
        # 有序列表（支持缩进级别）
        if re.match(r'^[\s]*\d+\.\s', line):
            items = []
            while i < len(lines) and re.match(r'^[\s]*\d+\.\s', lines[i]):
                match = re.match(r'^(\s*)(\d+)\.\s+(.*)$', lines[i])
                if match:
                    indent = len(match.group(1))
                    level = indent // 2  # 每2个空格为一级
                    text = match.group(3)
                    items.append({'level': level, 'text': text})
                i += 1
            blocks.append(BlockElement('list', items, level=1))  # level=1表示有序列表
            continue
        
        # 水平线
        if re.match(r'^[\s]*[-\*_]{3,}[\s]*$', line):
            blocks.append(BlockElement('hr', ''))
            i += 1
            continue
        
        # 图片（单独一行）
        img_match = re.match(r'^!\[([^\]]*)\]\(([^\)]+)\)$', line.strip())
        if img_match:
            blocks.append(BlockElement('image', img_match.group(1), language=img_match.group(2)))
            i += 1
            continue
        
        # 普通段落
        para_lines = []
        while i < len(lines) and lines[i].strip():
            # 遇到特殊块元素则停止
            if lines[i].startswith('#') or lines[i].strip().startswith('```') or \
               lines[i].strip().startswith('$$') or lines[i].startswith('>') or \
               re.match(r'^[\s]*[\*\-\+]\s', lines[i]) or re.match(r'^[\s]*\d+\.\s', lines[i]) or \
               re.match(r'^[\s]*[-\*_]{3,}[\s]*$', lines[i]):
                break
            # 检查是否是表格
            if '|' in lines[i] and i + 1 < len(lines) and re.match(r'^[\s\|\:\-]+$', lines[i + 1]):
                break
            para_lines.append(lines[i])
            i += 1
        
        if para_lines:
            blocks.append(BlockElement('paragraph', ' '.join(para_lines)))
    
    return blocks


def _split_table_cells(line: str) -> List[str]:
    """
    按 | 分割表格单元格，正确处理转义的竖线 (\\|)
    
    Args:
        line: 表格行文本
        
    Returns:
        单元格列表
    """
    cells = []
    current_cell = []
    i = 0
    
    while i < len(line):
        if i > 0 and line[i] == '|' and line[i - 1] == '\\':
            # 转义的竖线，替换前一个反斜杠并添加竖线
            current_cell[-1] = '|'
            i += 1
        elif line[i] == '|':
            # 未转义的竖线，分割单元格
            cells.append(''.join(current_cell).strip())
            current_cell = []
            i += 1
        else:
            current_cell.append(line[i])
            i += 1
    
    # 添加最后一个单元格
    if current_cell or cells:
        cells.append(''.join(current_cell).strip())
    
    return cells


def parse_table(table_text: str) -> Tuple[List[str], List[List[str]], List[str]]:
    """
    解析表格文本，支持转义竖线 (\\|)
    返回 (headers, rows, alignments)
    """
    lines = [l.strip() for l in table_text.strip().split('\n') if l.strip()]
    if len(lines) < 2:
        return [], [], []
    
    # 解析表头（使用新的分割方法）
    headers = _split_table_cells(lines[0])
    # 移除首尾的空元素（如果行是 |a|b| 格式）
    if headers and headers[0] == '':
        headers = headers[1:]
    if headers and headers[-1] == '':
        headers = headers[:-1]
    
    # 解析对齐方式
    align_cells = _split_table_cells(lines[1])
    if align_cells and align_cells[0] == '':
        align_cells = align_cells[1:]
    if align_cells and align_cells[-1] == '':
        align_cells = align_cells[:-1]
    
    alignments = []
    for cell in align_cells:
        cell = cell.strip()
        if cell.startswith(':') and cell.endswith(':'):
            alignments.append('center')
        elif cell.endswith(':'):
            alignments.append('right')
        else:
            alignments.append('left')
    
    # 解析数据行
    rows = []
    for line in lines[2:]:
        row = _split_table_cells(line)
        # 移除首尾空元素
        if row and row[0] == '':
            row = row[1:]
        if row and row[-1] == '':
            row = row[:-1]
        if row:
            rows.append(row)
    
    return headers, rows, alignments
