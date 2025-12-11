# -*- coding: utf-8 -*-
"""
工具函数模块 - 图片下载、路径处理等
"""

import os
import re
import tempfile
import requests
from urllib.parse import urlparse
from PIL import Image
from io import BytesIO


def is_url(path: str) -> bool:
    """判断路径是否为URL"""
    return path.startswith(('http://', 'https://'))


def download_image(url: str, timeout: int = 30) -> str:
    """
    下载网络图片到临时文件
    返回本地临时文件路径
    """
    try:
        response = requests.get(url, timeout=timeout)
        response.raise_for_status()
        
        # 从URL或Content-Type获取扩展名
        ext = get_image_extension(url, response.headers.get('Content-Type', ''))
        
        # 创建临时文件
        fd, temp_path = tempfile.mkstemp(suffix=ext)
        os.close(fd)
        
        with open(temp_path, 'wb') as f:
            f.write(response.content)
        
        return temp_path
    except Exception as e:
        print(f"图片下载失败: {url}, 错误: {e}")
        return None


def get_image_extension(url: str, content_type: str) -> str:
    """获取图片扩展名"""
    # 从Content-Type推断
    type_map = {
        'image/png': '.png',
        'image/jpeg': '.jpg',
        'image/gif': '.gif',
        'image/webp': '.webp',
        'image/bmp': '.bmp',
    }
    
    if content_type in type_map:
        return type_map[content_type]
    
    # 从URL推断
    parsed = urlparse(url)
    path = parsed.path.lower()
    for ext in ['.png', '.jpg', '.jpeg', '.gif', '.webp', '.bmp']:
        if path.endswith(ext):
            return ext
    
    return '.png'  # 默认PNG


def resolve_image_path(image_path: str, base_dir: str = None) -> str:
    """
    解析图片路径，支持相对路径和绝对路径
    """
    if is_url(image_path):
        return download_image(image_path)
    
    # 处理本地路径
    if os.path.isabs(image_path):
        return image_path if os.path.exists(image_path) else None
    
    # 相对路径
    if base_dir:
        full_path = os.path.join(base_dir, image_path)
        if os.path.exists(full_path):
            return full_path
    
    return image_path if os.path.exists(image_path) else None


def get_image_dimensions(image_path: str, max_width_inches: float = 6.0) -> tuple:
    """
    获取图片尺寸，自动缩放以适应页面宽度
    返回 (width_inches, height_inches)
    """
    try:
        with Image.open(image_path) as img:
            width_px, height_px = img.size
            dpi = img.info.get('dpi', (96, 96))
            dpi_x = dpi[0] if isinstance(dpi, tuple) else dpi
            
            # 计算原始英寸尺寸
            width_inches = width_px / dpi_x
            height_inches = height_px / dpi_x
            
            # 如果超过最大宽度，等比缩放
            if width_inches > max_width_inches:
                scale = max_width_inches / width_inches
                width_inches = max_width_inches
                height_inches *= scale
            
            return (width_inches, height_inches)
    except Exception as e:
        print(f"获取图片尺寸失败: {e}")
        return (4.0, 3.0)  # 默认尺寸


def clean_text(text: str) -> str:
    """清理文本中的特殊字符"""
    # 移除零宽字符
    text = re.sub(r'[\u200b\u200c\u200d\ufeff]', '', text)
    return text


def normalize_markdown(md_text: str) -> str:
    """
    规范化 Markdown 文本格式，确保元素之间有适当的空行
    
    主要处理以下问题：
    1. 标题前后缺少空行（如智谱清言等AI输出）
    2. 代码块前后缺少空行
    3. 列表、引用等块级元素前后缺少空行
    4. 表格前后缺少空行
    5. 清理多余的连续空行
    """
    # 统一换行符为 \n
    text = md_text.replace('\r\n', '\n').replace('\r', '\n')
    lines = text.split('\n')
    
    result = []
    in_code_block = False
    in_table = False
    prev_line_type = 'start'  # start, empty, text, heading, code, table, list, quote, hr
    
    for i, line in enumerate(lines):
        current_type = _get_line_type(line, in_code_block, in_table)
        
        # 代码块状态切换
        if line.startswith('```'):
            in_code_block = not in_code_block
        
        # 表格状态检测
        if current_type == 'table':
            in_table = True
        elif in_table and current_type not in ('table', 'empty'):
            in_table = False
        
        # 决定是否需要在当前行前添加空行
        need_blank_before = _should_add_blank_line(prev_line_type, current_type)
        
        if need_blank_before and result and result[-1].strip():
            result.append('')
        
        result.append(line)
        
        # 决定是否需要在当前行后添加空行
        need_blank_after = _should_add_blank_after(current_type, i, lines)
        
        if need_blank_after and i + 1 < len(lines) and lines[i + 1].strip():
            result.append('')
        
        # 更新前一行类型
        prev_line_type = current_type if line.strip() else 'empty'
    
    text = '\n'.join(result)
    
    # 清理多余的连续空行（超过2个连续换行符压缩为2个）
    text = re.sub(r'\n{3,}', '\n\n', text)
    
    return text


def _get_line_type(line: str, in_code_block: bool, in_table: bool) -> str:
    """判断行的类型"""
    stripped = line.strip()
    
    if not stripped:
        return 'empty'
    
    if in_code_block:
        return 'code'
    
    # 代码块边界
    if line.startswith('```'):
        return 'code'
    
    # 标题
    if re.match(r'^#{1,6}\s+', line):
        return 'heading'
    
    # 表格
    if line.startswith('|') or (line.endswith('|') and '|' in line):
        return 'table'
    
    # 分隔线
    if re.match(r'^[-*_]{3,}$', stripped):
        return 'hr'
    
    # 列表（无序）
    if re.match(r'^[-*+]\s', line):
        return 'list'
    
    # 列表（有序）
    if re.match(r'^\d+\.\s', line):
        return 'list'
    
    # 块级公式
    if stripped.startswith('$$'):
        return 'math'
    
    # 引用
    if line.startswith('>'):
        return 'quote'
    
    return 'text'


def _should_add_blank_line(prev_type: str, current_type: str) -> bool:
    """判断两种类型之间是否需要空行"""
    # 文档开头或前一行是空行，不需要
    if prev_type in ('start', 'empty'):
        return False
    
    # 当前行是空行，不需要
    if current_type == 'empty':
        return False
    
    # 标题前需要空行（除非前面是标题）
    if current_type == 'heading':
        return prev_type not in ('heading',)
    
    # 代码块前需要空行
    if current_type == 'code' and prev_type != 'code':
        return True
    
    # 表格前需要空行
    if current_type == 'table' and prev_type not in ('table',):
        return True
    
    # 列表前需要空行（除非前面是列表）
    if current_type == 'list' and prev_type not in ('list',):
        return True
    
    # 块级公式前需要空行
    if current_type == 'math' and prev_type not in ('math',):
        return True
    
    # 引用前需要空行
    if current_type == 'quote' and prev_type not in ('quote',):
        return True
    
    # 分隔线前需要空行
    if current_type == 'hr':
        return True
    
    return False


def _should_add_blank_after(current_type: str, index: int, lines: list) -> bool:
    """判断当前行后是否需要空行"""
    # 最后一行不需要
    if index >= len(lines) - 1:
        return False
    
    next_line = lines[index + 1].strip()
    
    # 下一行已经是空行，不需要
    if not next_line:
        return False
    
    # 标题后需要空行
    if current_type == 'heading':
        return True
    
    # 代码块结束后需要空行
    if current_type == 'code' and lines[index].startswith('```'):
        # 检查是结束标记
        in_code = False
        for i in range(index):
            if lines[i].startswith('```'):
                in_code = not in_code
        if not in_code:  # 结束标记
            return True
    
    # 分隔线后需要空行
    if current_type == 'hr':
        return True
    
    return False


def convert_latex_delimiters(text: str) -> str:
    """
    将多种 LaTeX 公式格式统一转换为 $...$ 和 $$...$$ 格式
    
    支持转换：
    - \\[...\\] -> $$...$$  (块级公式)
    - \\(...\\) -> $...$    (行内公式)
    """
    # 块级公式 \[...\]
    text = re.sub(r'\\\[(.+?)\\\]', r'$$\1$$', text, flags=re.DOTALL)
    
    # 行内公式 \(...\)
    text = re.sub(r'\\\((.+?)\\\)', r'$\1$', text, flags=re.DOTALL)
    
    return text


def extract_latex_from_text(text: str) -> list:
    """
    从文本中提取LaTeX公式 - 改进版，更准确识别
    返回 [(start, end, formula, is_block), ...]
    """
    results = []
    
    # 块级公式 $$...$$ (支持多行)
    for match in re.finditer(r'\$\$(.+?)\$\$', text, re.DOTALL):
        formula = match.group(1).strip()
        if formula:  # 忽略空公式
            results.append((match.start(), match.end(), formula, True))
    
    # 行内公式 $...$ - 改进匹配规则
    # 要求：1. 公式内容不能为空或仅空白
    #       2. 公式内容必须包含数学字符或LaTeX命令
    #       3. 避免匹配单个普通字符如 $y$
    inline_pattern = r'(?<!\$)\$(?!\$)([^\$]+?)(?<!\$)\$(?!\$)'
    
    for match in re.finditer(inline_pattern, text):
        formula = match.group(1).strip()
        
        # 跳过空公式
        if not formula:
            continue
        
        # 检查是否与块级公式重叠
        overlap = False
        for start, end, _, _ in results:
            if match.start() >= start and match.end() <= end:
                overlap = True
                break
        if overlap:
            continue
        
        # 判断是否为有效的数学公式（而非普通文字）
        if is_valid_math_formula(formula):
            results.append((match.start(), match.end(), formula, False))
    
    # 按位置排序
    results.sort(key=lambda x: x[0])
    return results


def is_valid_math_formula(formula: str) -> bool:
    """
    判断是否为有效的数学公式
    单个字母变量（如 r, x, y）也是有效公式，应以斜体显示
    """
    formula = formula.strip()
    
    # 空公式无效
    if not formula:
        return False
    
    # 包含LaTeX命令的一定是公式
    if '\\' in formula:
        return True
    
    # 包含数学运算符的是公式
    math_operators = ['+', '-', '*', '/', '=', '<', '>', '^', '_', '{', '}', 
                      '×', '÷', '±', '∞', '∑', '∫', '∏', '√']
    if any(op in formula for op in math_operators):
        return True
    
    # 包含希腊字母的是公式
    greek_letters = 'αβγδεζηθικλμνξπρστυφχψω'
    if any(c in greek_letters for c in formula.lower()):
        return True
    
    # 单个或多个拉丁字母变量是有效公式（如 r, x, y, Cin, Cout）
    if re.match(r'^[a-zA-Z]+$', formula):
        return True
    
    # 包含数字和字母组合（如 x2, a1）
    if re.search(r'[a-zA-Z]\d|\d[a-zA-Z]', formula):
        return True
    
    # 纯数字是有效公式（如常数）
    if re.match(r'^\d+(\.\d+)?$', formula):
        return True
    
    # 如果是普通文字（包含空格、标点符号或中文字符），则不是公式
    # 检查是否包含中文字符或常见标点
    if re.search(r'[\u4e00-\u9fff]', formula):  # 包含中文
        return False
    if re.search(r'[,.，。!?;:]', formula):  # 包含标点符号
        return False
    if ' ' in formula and not any(c in formula for c in ['+', '-', '=', '<', '>']):
        # 包含空格但不是数学表达式
        return False
    
    # 其他情况默认不是公式
    return False
