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
    FOOTNOTE_REF = "footnote_ref" # 脚注引用 [^1]
    ENDNOTE_REF = "endnote_ref"   # 尾注引用 [^^1]


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
    line_start: int = 0  # 源码起始行号
    line_end: int = 0    # 源码结束行号


def parse_inline(text: str) -> List[InlineElement]:
    """
    解析行内元素，返回元素列表
    这是预览和Word导出共用的核心解析逻辑
    使用简单可靠的顺序匹配方法
    """
    elements = []
    
    # 简化的正则模式（按优先级排序，使用非贪婪匹配）
    # 合并成一个大正则，每个模式用括号分组
    # 注意：公式匹配应具有最高优先级，以避免与斜体等符号冲突
    pattern = (
        r'(!\[[^\]]*\]\([^\)]+\))'           # 1: 图片
        r'|(\[[^\]]+\]\([^\)]+\))'            # 2: 链接
        r'|(\$\$(?:\\\$|[^\$])+?\$\$)'         # 3: 块级公式（行内模式）
        r'|(?<!\\)\$((?:\\\$|[^\$])+?)(?<!\\)\$'  # 4: 行内公式（负向断言支持转义）
        r'|(`[^`]+`)'                          # 5: 行内代码
        r'|(<br\s*/?>)'                         # 6: 换行标签
        r'|(<sup>[^<]+</sup>)'                 # 7: 上标 HTML
        r'|(<sub>[^<]+</sub>)'                 # 8: 下标 HTML
        r'|(\*\*\*.+?\*\*\*)'                  # 9: 粗斜体
        r'|(___.+?___)'                        # 10: 粗斜体
        r'|(\*\*.+?\*\*)'                      # 11: 粗体
        r'|(__.+?__)'                          # 12: 粗体
        r'|(?<!\*)\*(?!\*)([^\*\s][^\*]*[^\*\s]|[^\*\s])\*(?!\*)'  # 13: 斜体
        r'|(?<!_)_(?!_)([^_\s][^_]*[^_\s]|[^_\s])_(?!_)'    # 14: 斜体
        r'|(~~.+?~~)'                          # 15: 删除线
        r'|(\[\^\^[^\]]+\])'                    # 16: 尾注引用
        r'|(\[\^[^\]]+\])'                      # 17: 脚注引用
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
        
        elif full_match.startswith('[^^'):
            # 尾注引用 [^^1] - 必须在脚注之前检查（更具体的模式优先）
            content = full_match[3:-1]  # 去掉 [^^ 和 ]
            elements.append(InlineElement(InlineType.ENDNOTE_REF, content))
        
        elif full_match.startswith('[^'):
            # 脚注引用 [^1] - 必须在普通链接之前检查
            content = full_match[2:-1]  # 去掉 [^ 和 ]
            elements.append(InlineElement(InlineType.FOOTNOTE_REF, content))
        
        elif full_match.startswith('['):
            # 链接 [text](url)
            m = re.match(r'\[([^\]]+)\]\(([^\)]+)\)', full_match)
            if m:
                elements.append(InlineElement(InlineType.LINK, m.group(1), m.group(2)))
        
        elif full_match.startswith('`'):
            # 行内代码 `code`
            content = full_match[1:-1]
            elements.append(InlineElement(InlineType.CODE, content))
        
        elif full_match.startswith('$$'):
            # 块级公式（出现在行内）
            content = full_match[2:-2]
            elements.append(InlineElement(InlineType.MATH, content))

        elif full_match.startswith('$'):
            # 行内公式 $formula$
            content = full_match[1:-1]
            # 简单校验，如果是转义的美元符号（单独一个且前面是反斜杠），不作为公式
            if content.startswith('\\$') and len(content) == 2:
                 elements.append(InlineElement(InlineType.TEXT, full_match))
            else:
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

        elif re.match(r'\[\^([^\]]+)\]', full_match):
            # 脚注引用 [^1]
            content = full_match[2:-1]
            elements.append(InlineElement(InlineType.FOOTNOTE_REF, content))
        
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
        start_idx = i + 1
        
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
            blocks.append(BlockElement('code_block', '\n'.join(code_lines), language=lang, line_start=start_idx, line_end=i))
            continue
        
        # 块级公式 $$ 或 \begin{...} 或 \[...\]
        stripped_line = line.strip()
        if stripped_line.startswith('$$') or stripped_line.startswith('\\begin{') or stripped_line.startswith('\\['):
            # 1. 处理 \begin{...} 环境 (最高优先级，严禁截断)
            if stripped_line.startswith('\\begin{'):
                env_match = re.match(r'^\\begin\{([^}*]+)\*?\}', stripped_line)
                if env_match:
                    env_name = env_match.group(1)
                    formula_lines = []
                    depth = 0
                    while i < len(lines):
                        line_content = lines[i]
                        formula_lines.append(line_content)
                        # 检查开始标签（含带星号版本）
                        if f'\\begin{{{env_name}}}' in line_content or f'\\begin{{{env_name}*}}' in line_content:
                            depth += 1
                        # 检查结束标签
                        if f'\\end{{{env_name}}}' in line_content or f'\\end{{{env_name}*}}' in line_content:
                            depth -= 1
                            if depth <= 0:
                                i += 1
                                break
                        i += 1
                    # 关键修复：将 align 等块标记为特定的 math_block_env 类型以便预览区特殊处理
                    blocks.append(BlockElement('math_block', '\n'.join(formula_lines), language=env_name, line_start=start_idx, line_end=i))
                    continue

            # 2. 处理 \[ ... \] 环境
            if stripped_line.startswith('\\['):
                formula_lines = []
                while i < len(lines):
                    line_content = lines[i]
                    formula_lines.append(line_content)
                    if '\\]' in line_content:
                        i += 1
                        break
                    i += 1
                blocks.append(BlockElement('math_block', '\n'.join(formula_lines), line_start=start_idx, line_end=i))
                continue

            # 3. 处理 $$ 公式
            # 检查是否是单行 $$公式$$
            single_match = re.match(r'^\$\$(.+?)\$\$$', stripped_line)
            if single_match:
                formula = single_match.group(1).strip()
                blocks.append(BlockElement('math_block', formula, line_start=start_idx, line_end=start_idx))
                i += 1
                continue
            
            # 检查是否是多行 $$ 开头
            if stripped_line == '$$' or stripped_line.startswith('$$'):
                formula_lines = []
                # 如果第一行 $$ 后面还有内容
                if len(stripped_line) > 2:
                    formula_lines.append(line[line.find('$$')+2:])
                
                i += 1
                while i < len(lines):
                    if lines[i].strip().endswith('$$'):
                        last_line = lines[i].strip()
                        if last_line != '$$':
                            formula_lines.append(lines[i][:lines[i].find('$$')])
                        i += 1
                        break
                    formula_lines.append(lines[i])
                    i += 1
                blocks.append(BlockElement('math_block', '\n'.join(formula_lines), line_start=start_idx, line_end=i))
                continue
        
        # 标题 #
        if line.startswith('#'):
            level = len(re.match(r'^#+', line).group())
            content = line[level:].strip()
            blocks.append(BlockElement('heading', content, level=min(level, 4), line_start=start_idx, line_end=start_idx))
            i += 1
            continue
        
        # 表格（对分隔行进行 strip 处理，确保匹配）
        if '|' in line and i + 1 < len(lines) and re.match(r'^[\s\|\:\-]+$', lines[i + 1].strip()):
            table_lines = []
            while i < len(lines) and '|' in lines[i]:
                table_lines.append(lines[i])
                i += 1
            blocks.append(BlockElement('table', '\n'.join(table_lines), line_start=start_idx, line_end=i))
            continue
        
        # 引用 >
        if line.startswith('>'):
            quote_lines = []
            while i < len(lines) and lines[i].startswith('>'):
                quote_lines.append(lines[i][1:].strip())
                i += 1
            blocks.append(BlockElement('quote', '\n'.join(quote_lines), line_start=start_idx, line_end=i))
            continue
        
        # 无序列表（包括任务列表）
        if re.match(r'^[\s]*[\*\-\+]\s', line):
            items = []
            while i < len(lines) and re.match(r'^[\s]*[\*\-\+]\s', lines[i]):
                current_line_idx = i + 1
                item_text = re.sub(r'^[\s]*[\*\-\+]\s*', '', lines[i])
                # 检查是否是任务列表
                task_match = re.match(r'^\[([ xX])\]\s*(.*)', item_text)
                if task_match:
                    checked = task_match.group(1).lower() == 'x'
                    text = task_match.group(2)
                    items.append({'type': 'task', 'checked': checked, 'text': text, 'line': current_line_idx})
                else:
                    items.append({'type': 'item', 'text': item_text, 'line': current_line_idx})
                i += 1
            blocks.append(BlockElement('list', items, level=0, line_start=start_idx, line_end=i))
            continue
        
        # 有序列表（支持缩进级别）
        if re.match(r'^[\s]*\d+\.\s', line):
            items = []
            while i < len(lines) and re.match(r'^[\s]*\d+\.\s', lines[i]):
                current_line_idx = i + 1
                match = re.match(r'^(\s*)(\d+)\.\s+(.*)$', lines[i])
                if match:
                    indent = len(match.group(1))
                    level = indent // 2  # 每2个空格为一级
                    text = match.group(3)
                    items.append({'level': level, 'text': text, 'line': current_line_idx})
                i += 1
            blocks.append(BlockElement('list', items, level=1, line_start=start_idx, line_end=i))  # level=1表示有序列表
            continue
        
        # 水平线
        if re.match(r'^[\s]*[-\*_]{3,}[\s]*$', line):
            blocks.append(BlockElement('hr', '', line_start=start_idx, line_end=start_idx))
            i += 1
            continue
        
        # 图片（单独一行）
        img_match = re.match(r'^!\[([^\]]*)\]\(([^\)]+)\)$', line.strip())
        if img_match:
            blocks.append(BlockElement('image', img_match.group(1), language=img_match.group(2), line_start=start_idx, line_end=start_idx))
            i += 1
            continue

        # 脚注定义 [^1]: 内容
        footnote_match = re.match(r'^\s*\[\^([^\]]+)\][:：]\s*(.*)', line.strip())
        if footnote_match:
            ref = footnote_match.group(1)
            content = footnote_match.group(2)
            blocks.append(BlockElement('footnote_def', content, language=ref, line_start=start_idx, line_end=start_idx))
            i += 1
            continue
        
        # 普通段落
        para_lines = []
        while i < len(lines) and lines[i].strip():
            # 遇到特殊块元素则停止
            if lines[i].startswith('#') or lines[i].strip().startswith('```') or \
               lines[i].strip().startswith('$$') or lines[i].strip().startswith('\\begin{') or \
               lines[i].strip().startswith('\\[') or lines[i].startswith('>') or \
               re.match(r'^[\s]*[\*\-\+]\s', lines[i]) or re.match(r'^[\s]*\d+\.\s', lines[i]) or \
               re.match(r'^[\s]*[-\*_]{3,}[\s]*$', lines[i]):
                break
            # 检查是否是表格
            if '|' in lines[i] and i + 1 < len(lines) and re.match(r'^[\s\|\:\-]+$', lines[i + 1]):
                break
            para_lines.append(lines[i])
            i += 1
        
        if para_lines:
            blocks.append(BlockElement('paragraph', ' '.join(para_lines), line_start=start_idx, line_end=i))
    
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
