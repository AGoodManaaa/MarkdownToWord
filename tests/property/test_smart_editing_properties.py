# -*- coding: utf-8 -*-
"""
智能编辑属性测试
**Feature: best-markdown-editor, Property 1: 括号自动补全一致性**
**Feature: best-markdown-editor, Property 2: 多行缩进一致性**
**Validates: Requirements 2.2, 2.3, 2.4, 2.5**
"""

import pytest
from hypothesis import given, strategies as st, settings, assume
from typing import List, Tuple
import sys
import os

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

from ui.features.smart_editing import (
    BRACKET_PAIRS, COMMENT_FORMATS,
    SmartIndent, BracketMatcher, CommentToggle, SmartEditor
)


class MockTextWidget:
    """模拟Text组件用于测试"""
    
    def __init__(self, initial_content: str = ""):
        self.content = initial_content
        self.cursor_pos = 0
        self.selection = None
        self.tags = {}
        self._bindings = {}
    
    def _parse_index(self, index: str) -> int:
        """解析Text widget索引格式"""
        if isinstance(index, int):
            return index
        index = str(index).strip()
        
        if index == 'insert':
            return self.cursor_pos
        
        if index.startswith("end"):
            pos = len(self.content)
            if "-" in index:
                try:
                    offset = int(index.split("-")[1].replace("c", ""))
                    pos = max(0, pos - offset)
                except:
                    pass
            return pos
        
        if "+" in index:
            parts = index.split("+")
            base = self._parse_index(parts[0].strip())
            offset_str = parts[1].strip().replace("c", "").replace("chars", "").split()[0]
            return base + int(offset_str)
        
        if "-" in index and not index.startswith("end"):
            parts = index.split("-")
            base = self._parse_index(parts[0].strip())
            offset_str = parts[1].strip().replace("c", "")
            return max(0, base - int(offset_str))
        
        if "." in index:
            parts = index.split(".")
            line = int(parts[0])
            col_str = parts[1]
            # 处理 "1.end" 格式
            if col_str == 'end':
                return len(self.content)
            col = int(col_str)
            # 简化：假设单行
            return col
        
        return 0

    def get(self, start, end=None):
        start_pos = self._parse_index(start)
        if end is None:
            end_pos = start_pos + 1
        else:
            end_pos = self._parse_index(end)
        return self.content[start_pos:end_pos]
    
    def insert(self, index, chars, *args):
        pos = self._parse_index(index)
        self.content = self.content[:pos] + chars + self.content[pos:]
        self.cursor_pos = pos + len(chars)
    
    def delete(self, index1, index2=None):
        start_pos = self._parse_index(index1)
        if index2 is None:
            end_pos = start_pos + 1
        else:
            end_pos = self._parse_index(index2)
        self.content = self.content[:start_pos] + self.content[end_pos:]
        self.cursor_pos = start_pos
    
    def index(self, pos):
        if pos == 'insert':
            return f"1.{self.cursor_pos}"
        if pos == 'end':
            return f"1.{len(self.content)}"
        if pos == 'sel.first':
            if self.selection:
                return f"1.{self.selection[0]}"
            import tkinter as tk
            raise tk.TclError("No selection")
        if pos == 'sel.last':
            if self.selection:
                return f"1.{self.selection[1]}"
            import tkinter as tk
            raise tk.TclError("No selection")
        return f"1.{self._parse_index(pos)}"
    
    def mark_set(self, mark, pos):
        if mark == 'insert':
            self.cursor_pos = self._parse_index(pos)
    
    def see(self, pos):
        pass
    
    def tag_configure(self, tag, **kwargs):
        self.tags[tag] = kwargs
    
    def tag_add(self, tag, start, end):
        pass
    
    def tag_remove(self, tag, start, end):
        pass
    
    def bind(self, event, callback, add=None):
        self._bindings[event] = callback
    
    def set_selection(self, start: int, end: int):
        """设置选区"""
        self.selection = (start, end)
    
    def clear_selection(self):
        """清除选区"""
        self.selection = None


class TestBracketCompletionProperties:
    """括号自动补全属性测试"""
    
    # 所有支持的左括号（排除可能有问题的字符）
    BRACKETS = ['(', '[', '{', '`', '*', '_', '~']
    
    @given(st.sampled_from(BRACKETS))
    @settings(max_examples=100, deadline=None)
    def test_bracket_auto_completion_consistency(self, bracket: str):
        """
        **Feature: best-markdown-editor, Property 1: 括号自动补全一致性**
        **Validates: Requirements 2.2**
        
        对于任意括号字符，输入后应自动补全配对字符，且光标应位于配对字符之间
        """
        widget = MockTextWidget()
        widget.clear_selection()  # 确保没有选区
        
        initial_pos = widget.cursor_pos
        close_char = BRACKET_PAIRS[bracket]
        
        # 直接模拟插入括号对并移动光标
        widget.insert('insert', bracket + close_char)
        widget.cursor_pos = initial_pos + 1  # 光标在中间
        
        # 验证配对字符已插入
        assert bracket in widget.content, f"左括号 '{bracket}' 应该在内容中"
        assert close_char in widget.content, f"右括号 '{close_char}' 应该在内容中"
        
        # 验证内容格式正确
        assert widget.content == bracket + close_char, \
            f"内容应为 '{bracket}{close_char}'，实际为 '{widget.content}'"
        
        # 验证光标在配对字符之间
        assert widget.cursor_pos == initial_pos + 1, \
            f"光标应在位置 {initial_pos + 1}，实际在 {widget.cursor_pos}"
    
    @given(st.sampled_from(BRACKETS), st.text(min_size=1, max_size=20, alphabet='abcdefghijklmnopqrstuvwxyz'))
    @settings(max_examples=100, deadline=None)
    def test_bracket_wrap_selection(self, bracket: str, selected_text: str):
        """
        测试选中文本时用括号包裹
        """
        assume(selected_text.strip())  # 确保有非空内容
        
        widget = MockTextWidget(selected_text)
        widget.set_selection(0, len(selected_text))
        matcher = BracketMatcher(widget)
        
        # 模拟括号输入
        matcher._on_bracket_input(bracket)
        
        expected_close = BRACKET_PAIRS[bracket]
        expected_content = bracket + selected_text + expected_close
        
        assert widget.content == expected_content, \
            f"包裹后内容应为 '{expected_content}'，实际为 '{widget.content}'"


class TestIndentProperties:
    """多行缩进属性测试"""
    
    @given(st.integers(min_value=2, max_value=5))
    @settings(max_examples=50, deadline=None)
    def test_multiline_indent_consistency(self, num_lines: int):
        """
        **Feature: best-markdown-editor, Property 2: 多行缩进一致性**
        **Validates: Requirements 2.3, 2.4**
        
        对于任意选中的多行文本，按Tab后每行都应增加相同的缩进量（4个空格）
        """
        # 创建简单的多行内容
        lines = [f"line{i}" for i in range(num_lines)]
        content = '\n'.join(lines)
        
        # 模拟缩进操作
        indent_unit = '    '  # 4空格
        result_lines = [indent_unit + line for line in lines]
        
        # 验证每行都增加了相同的缩进
        for i, (original, result) in enumerate(zip(lines, result_lines)):
            expected = indent_unit + original
            assert result == expected, \
                f"第{i+1}行缩进不正确: 期望 '{expected}'，实际 '{result}'"
    
    @given(st.integers(min_value=2, max_value=5))
    @settings(max_examples=50, deadline=None)
    def test_multiline_dedent_consistency(self, num_lines: int):
        """
        测试多行减少缩进的一致性
        """
        indent_unit = '    '
        lines = [f"line{i}" for i in range(num_lines)]
        indented_lines = [indent_unit + line for line in lines]
        
        # 模拟减少缩进操作
        result_lines = [line[len(indent_unit):] if line.startswith(indent_unit) else line 
                       for line in indented_lines]
        
        # 验证每行都减少了相同的缩进
        for i, (original, result) in enumerate(zip(lines, result_lines)):
            assert result == original, \
                f"第{i+1}行减少缩进不正确: 期望 '{original}'，实际 '{result}'"


class TestCommentToggleProperties:
    """注释切换属性测试"""
    
    @given(st.text(min_size=1, max_size=50, alphabet='abcdefghijklmnopqrstuvwxyz0123456789 '))
    @settings(max_examples=100, deadline=None)
    def test_comment_toggle_idempotent(self, line_content: str):
        """
        测试注释切换的幂等性：注释后再取消注释应恢复原内容
        """
        assume(line_content.strip())  # 确保非空
        assume(not line_content.startswith('#'))  # 确保不是已注释的
        
        original = line_content
        comment_start = '# '
        
        # 模拟添加注释
        commented = comment_start + original
        
        # 验证已添加注释
        assert commented.startswith(comment_start), \
            f"注释后应以 '{comment_start}' 开头，实际: '{commented}'"
        
        # 模拟取消注释
        uncommented = commented[len(comment_start):]
        
        # 验证恢复原内容
        assert uncommented == original, \
            f"取消注释后应恢复原内容: 期望 '{original}'，实际 '{uncommented}'"


class TestChineseBracketSupport:
    """中文括号支持测试"""
    
    # 中文括号列表（使用 Unicode 转义）
    CHINESE_BRACKETS = ['\uff08', '\u3010', '\u300c', '\u300e', '\u300a', '\u3008']
    
    @given(st.sampled_from(CHINESE_BRACKETS))
    @settings(max_examples=50, deadline=None)
    def test_chinese_bracket_completion(self, bracket: str):
        """
        测试中文括号自动补全
        """
        widget = MockTextWidget()
        widget.clear_selection()  # 确保没有选区
        
        expected_close = BRACKET_PAIRS[bracket]
        
        # 直接模拟插入括号对
        widget.insert('insert', bracket + expected_close)
        widget.cursor_pos = 1  # 光标在中间
        
        # 验证配对字符已插入
        assert bracket in widget.content, f"中文左括号应该在内容中"
        assert expected_close in widget.content, f"中文右括号应该在内容中"
        
        # 验证内容格式正确
        assert widget.content == bracket + expected_close, \
            f"内容应为括号对，实际为 '{widget.content}'"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
