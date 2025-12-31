# -*- coding: utf-8 -*-
"""Markdown 生成器模块 - 将 OCR 结果转换为 Markdown"""

import re
from typing import List, Optional
from .ocr_engine import OCRResult, OCRRegion, ContentType


class MarkdownGenerator:
    """将 OCR 结果转换为 Markdown"""
    
    def __init__(self):
        self.indent = "  "
        self.list_markers = ['-', '*', '•', '·', '○', '●', '◆', '◇', '►', '▸']
        self.number_pattern = re.compile(r'^(\d+)[.、)）]\s*')
    
    def generate(self, result: OCRResult, include_source_comment: bool = True) -> str:
        """生成完整的 Markdown 文档
        
        Args:
            result: OCR 识别结果
            include_source_comment: 是否包含源图片注释
            
        Returns:
            Markdown 格式的文本
        """
        if not result.has_content:
            return ""
        
        parts = []
        
        # 添加源图片注释
        if include_source_comment and result.source_image:
            parts.append(f"<!-- OCR Source: {result.source_image} -->")
            parts.append("")
        
        # 按区域生成内容
        for region in result.regions:
            md_content = self._region_to_markdown(region)
            if md_content:
                parts.append(md_content)
                parts.append("")  # 区域间空行
        
        return '\n'.join(parts).strip()
    
    def _region_to_markdown(self, region: OCRRegion) -> str:
        """将单个区域转换为 Markdown
        
        Args:
            region: OCR 识别区域
            
        Returns:
            Markdown 文本
        """
        if region.content_type == ContentType.TEXT:
            return self.text_to_markdown(region)
        elif region.content_type == ContentType.TABLE:
            return self.table_to_markdown(region)
        elif region.content_type == ContentType.FORMULA:
            return self.formula_to_markdown(region)
        else:
            return region.content
    
    def text_to_markdown(self, region: OCRRegion) -> str:
        """文字转 Markdown
        
        Args:
            region: 文字区域
            
        Returns:
            Markdown 文本
        """
        text = region.content.strip()
        if not text:
            return ""
        
        # 检测并转换标题
        text = self._detect_heading_structure(text)
        
        # 检测并转换列表
        text = self._detect_list_structure(text)
        
        return text
    
    def table_to_markdown(self, region: OCRRegion) -> str:
        """表格转 Markdown
        
        Args:
            region: 表格区域
            
        Returns:
            Markdown 表格
        """
        content = region.content.strip()
        
        # 如果已经是 Markdown 表格格式，直接返回
        if content.startswith('|') and '|' in content:
            return self._validate_markdown_table(content)
        
        # 尝试从原始数据构建表格
        if region.raw_data and 'cells' in region.raw_data:
            return self._build_table_from_cells(region.raw_data['cells'])
        
        return content
    
    def formula_to_markdown(self, region: OCRRegion, inline: bool = False) -> str:
        """公式转 Markdown LaTeX
        
        Args:
            region: 公式区域
            inline: 是否为行内公式
            
        Returns:
            Markdown LaTeX 公式
        """
        latex = region.content.strip()
        if not latex:
            return ""
        
        # 清理 LaTeX 代码
        latex = self._clean_latex(latex)
        
        # 根据公式长度和复杂度决定是行内还是块级
        if inline or self._is_simple_formula(latex):
            return f"${latex}$"
        else:
            return f"$$\n{latex}\n$$"
    
    def _detect_heading_structure(self, text: str) -> str:
        """检测并转换标题结构
        
        Args:
            text: 原始文本
            
        Returns:
            转换后的文本
        """
        lines = text.split('\n')
        result_lines = []
        
        for line in lines:
            stripped = line.strip()
            
            # 检测中文标题格式（如：一、标题）
            chinese_heading = re.match(r'^([一二三四五六七八九十]+)[、.．]\s*(.+)$', stripped)
            if chinese_heading:
                level = len(chinese_heading.group(1))
                level = min(level, 3)  # 最多三级标题
                result_lines.append('#' * level + ' ' + chinese_heading.group(2))
                continue
            
            # 检测数字标题格式（如：1. 标题 或 1.1 标题）
            num_heading = re.match(r'^(\d+(?:\.\d+)*)[.、)）]\s*(.+)$', stripped)
            if num_heading:
                num_parts = num_heading.group(1).split('.')
                level = min(len(num_parts), 6)
                result_lines.append('#' * level + ' ' + num_heading.group(2))
                continue
            
            result_lines.append(line)
        
        return '\n'.join(result_lines)
    
    def _detect_list_structure(self, text: str) -> str:
        """检测并转换列表结构
        
        Args:
            text: 原始文本
            
        Returns:
            转换后的文本
        """
        lines = text.split('\n')
        result_lines = []
        in_list = False
        
        for line in lines:
            stripped = line.strip()
            
            # 检测无序列表标记
            for marker in self.list_markers:
                if stripped.startswith(marker + ' ') or stripped.startswith(marker + '\t'):
                    result_lines.append('- ' + stripped[len(marker):].strip())
                    in_list = True
                    break
            else:
                # 检测有序列表
                num_match = self.number_pattern.match(stripped)
                if num_match:
                    content = stripped[num_match.end():].strip()
                    result_lines.append(f"{num_match.group(1)}. {content}")
                    in_list = True
                else:
                    # 普通行
                    if in_list and stripped:
                        # 列表后的非空行，添加空行分隔
                        if result_lines and not result_lines[-1] == '':
                            result_lines.append('')
                        in_list = False
                    result_lines.append(line)
        
        return '\n'.join(result_lines)
    
    def _validate_markdown_table(self, table: str) -> str:
        """验证并修复 Markdown 表格语法
        
        Args:
            table: Markdown 表格字符串
            
        Returns:
            修复后的表格
        """
        lines = table.strip().split('\n')
        if len(lines) < 2:
            return table
        
        # 解析表格
        rows = []
        for line in lines:
            if line.strip().startswith('|'):
                cells = [c.strip() for c in line.strip().strip('|').split('|')]
                rows.append(cells)
        
        if not rows:
            return table
        
        # 确定列数
        max_cols = max(len(row) for row in rows)
        
        # 重建表格
        result_lines = []
        for i, row in enumerate(rows):
            # 补齐列数
            while len(row) < max_cols:
                row.append('')
            
            result_lines.append('| ' + ' | '.join(row) + ' |')
            
            # 在第一行后添加分隔行
            if i == 0:
                result_lines.append('| ' + ' | '.join(['---'] * max_cols) + ' |')
        
        return '\n'.join(result_lines)
    
    def _build_table_from_cells(self, cells: List[List[str]]) -> str:
        """从单元格数据构建 Markdown 表格
        
        Args:
            cells: 二维单元格数据
            
        Returns:
            Markdown 表格
        """
        if not cells:
            return ""
        
        max_cols = max(len(row) for row in cells)
        
        lines = []
        for i, row in enumerate(cells):
            # 补齐列数
            while len(row) < max_cols:
                row.append('')
            
            lines.append('| ' + ' | '.join(str(c) for c in row) + ' |')
            
            if i == 0:
                lines.append('| ' + ' | '.join(['---'] * max_cols) + ' |')
        
        return '\n'.join(lines)
    
    def _clean_latex(self, latex: str) -> str:
        """清理 LaTeX 代码
        
        Args:
            latex: 原始 LaTeX
            
        Returns:
            清理后的 LaTeX
        """
        # 移除多余的空白
        latex = ' '.join(latex.split())
        
        # 移除已有的 $ 符号
        latex = latex.strip('$')
        
        # 修复常见问题
        latex = latex.replace('\\\\', '\\')  # 双反斜杠
        
        return latex.strip()
    
    def _is_simple_formula(self, latex: str) -> bool:
        """判断是否为简单公式（适合行内显示）
        
        Args:
            latex: LaTeX 代码
            
        Returns:
            是否为简单公式
        """
        # 简单公式的特征
        if len(latex) > 50:
            return False
        
        # 包含换行或对齐环境的是复杂公式
        complex_patterns = [
            r'\\begin\{',
            r'\\end\{',
            r'\\\\',
            r'\\newline',
            r'\\displaystyle',
        ]
        
        for pattern in complex_patterns:
            if re.search(pattern, latex):
                return False
        
        return True
    
    def merge_results(self, results: List[OCRResult]) -> str:
        """合并多个 OCR 结果
        
        Args:
            results: OCR 结果列表
            
        Returns:
            合并后的 Markdown
        """
        parts = []
        
        for i, result in enumerate(results):
            if result.has_content:
                md = self.generate(result, include_source_comment=True)
                if md:
                    if i > 0:
                        parts.append('\n---\n')  # 分隔符
                    parts.append(md)
        
        return '\n'.join(parts)
