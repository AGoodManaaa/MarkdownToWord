# -*- coding: utf-8 -*-
"""智能编辑功能模块 - 智能缩进、括号匹配、自动补全、注释切换"""

import re
import tkinter as tk
from typing import Optional, Tuple, List, Dict

try:
    import customtkinter as ctk
except ImportError:
    ctk = None


# 括号配对 - 包含中英文括号
# 使用 Unicode 转义避免编码问题
BRACKET_PAIRS: Dict[str, str] = {
    # 英文括号
    '(': ')',
    '[': ']',
    '{': '}',
    '"': '"',
    "'": "'",
    '`': '`',
    '*': '*',
    '_': '_',
    '~': '~',
    # 中文括号 (使用 Unicode 转义)
    '\uff08': '\uff09',  # （）
    '\u3010': '\u3011',  # 【】
    '\u300c': '\u300d',  # 「」
    '\u300e': '\u300f',  # 『』
    '\u300a': '\u300b',  # 《》
    '\u3008': '\u3009',  # 〈〉
    '\u201c': '\u201d',  # ""
    '\u2018': '\u2019',  # ''
}

# 代码块语言到注释格式的映射
COMMENT_FORMATS: Dict[str, Tuple[str, str]] = {
    'python': ('# ', ''),
    'javascript': ('// ', ''),
    'typescript': ('// ', ''),
    'java': ('// ', ''),
    'c': ('// ', ''),
    'cpp': ('// ', ''),
    'csharp': ('// ', ''),
    'go': ('// ', ''),
    'rust': ('// ', ''),
    'ruby': ('# ', ''),
    'php': ('// ', ''),
    'shell': ('# ', ''),
    'bash': ('# ', ''),
    'sql': ('-- ', ''),
    'html': ('<!-- ', ' -->'),
    'xml': ('<!-- ', ' -->'),
    'css': ('/* ', ' */'),
}


# 列表标记正则
LIST_PATTERNS = {
    'unordered': r'^(\s*)([-*+])\s+',
    'ordered': r'^(\s*)(\d+)\.\s+',
    'task': r'^(\s*)([-*+])\s+\[([ xX])\]\s+',
}


class SmartIndent:
    """智能缩进"""
    
    def __init__(self, text_widget):
        """
        初始化智能缩进
        
        Args:
            text_widget: tkinter Text 或 CTkTextbox 组件
        """
        self.text_widget = text_widget
        self._enabled = True
        self._indent_size = 4
        self._use_tabs = False
        
        # 获取底层 Text 组件
        if hasattr(text_widget, '_textbox'):
            self._text = text_widget._textbox
        else:
            self._text = text_widget
        
        # 绑定事件
        self._bind_events()
    
    def _bind_events(self):
        """绑定键盘事件"""
        self._text.bind('<Return>', self._on_return, add='+')
        self._text.bind('<Tab>', self._on_tab)
        self._text.bind('<Shift-Tab>', self._on_shift_tab)
        self._text.bind('<BackSpace>', self._on_backspace, add='+')
    
    def _get_current_line(self) -> Tuple[str, int]:
        """获取当前行内容和行号"""
        cursor = self._text.index('insert')
        line_num = int(cursor.split('.')[0])
        line_start = f"{line_num}.0"
        line_end = f"{line_num}.end"
        line_content = self._text.get(line_start, line_end)
        return line_content, line_num
    
    def _get_indent(self, line: str) -> str:
        """获取行首缩进"""
        match = re.match(r'^(\s*)', line)
        return match.group(1) if match else ''
    
    def _get_indent_unit(self) -> str:
        """获取缩进单位"""
        if self._use_tabs:
            return '\t'
        return ' ' * self._indent_size

    def _on_return(self, event=None):
        """回车键处理"""
        if not self._enabled:
            return
        
        line, line_num = self._get_current_line()
        cursor = self._text.index('insert')
        col = int(cursor.split('.')[1])
        
        # 获取光标前的内容
        line_before_cursor = line[:col]
        
        # 检查是否是列表项
        for pattern_name, pattern in LIST_PATTERNS.items():
            match = re.match(pattern, line_before_cursor)
            if match:
                return self._handle_list_return(match, pattern_name, line, col)
        
        # 普通缩进继承
        indent = self._get_indent(line)
        
        # 检查是否需要增加缩进（如代码块开始）
        if line_before_cursor.rstrip().endswith(':'):
            indent += self._get_indent_unit()
        
        # 插入换行和缩进
        self._text.insert('insert', '\n' + indent)
        return 'break'
    
    def _handle_list_return(self, match, pattern_name: str, line: str, col: int):
        """处理列表项回车"""
        indent = match.group(1)
        marker = match.group(2)
        
        # 检查列表项是否为空
        content_start = match.end()
        content = line[content_start:col].strip()
        
        if not content:
            # 空列表项，结束列表
            line_num = int(self._text.index('insert').split('.')[0])
            self._text.delete(f"{line_num}.0", f"{line_num}.end")
            indent_unit = self._get_indent_unit()
            new_indent = indent[:-len(indent_unit)] if indent and len(indent) >= len(indent_unit) else ''
            self._text.insert(f"{line_num}.0", new_indent)
            return 'break'
        
        # 生成新的列表标记
        if pattern_name == 'ordered':
            new_marker = str(int(marker) + 1) + '.'
        elif pattern_name == 'task':
            new_marker = marker + ' [ ]'
        else:
            new_marker = marker
        
        # 插入新列表项
        self._text.insert('insert', f'\n{indent}{new_marker} ')
        return 'break'

    def _on_tab(self, event=None):
        """Tab 键处理 - 支持多行缩进"""
        if not self._enabled:
            return
        
        # 检查是否有选中内容
        try:
            sel_start = self._text.index('sel.first')
            sel_end = self._text.index('sel.last')
            return self._indent_selection(sel_start, sel_end, increase=True)
        except tk.TclError:
            pass
        
        # 检查是否在代码块内
        if self._is_in_code_block():
            # 代码块内使用4空格
            self._text.insert('insert', '    ')
            return 'break'
        
        # 检查是否在列表项中
        line, line_num = self._get_current_line()
        for pattern in LIST_PATTERNS.values():
            match = re.match(pattern, line)
            if match:
                # 增加列表缩进
                indent_unit = self._get_indent_unit()
                self._text.insert(f"{line_num}.0", indent_unit)
                return 'break'
        
        # 普通 Tab 插入
        self._text.insert('insert', self._get_indent_unit())
        return 'break'
    
    def _on_shift_tab(self, event=None):
        """Shift+Tab 键处理 - 支持多行减少缩进"""
        if not self._enabled:
            return
        
        # 检查是否有选中内容
        try:
            sel_start = self._text.index('sel.first')
            sel_end = self._text.index('sel.last')
            return self._indent_selection(sel_start, sel_end, increase=False)
        except tk.TclError:
            pass
        
        # 减少当前行缩进
        line, line_num = self._get_current_line()
        indent = self._get_indent(line)
        indent_unit = self._get_indent_unit()
        
        if indent.endswith(indent_unit):
            self._text.delete(f"{line_num}.0", f"{line_num}.{len(indent_unit)}")
        elif indent:
            # 删除所有前导空白
            self._text.delete(f"{line_num}.0", f"{line_num}.{len(indent)}")
        
        return 'break'
    
    def _indent_selection(self, start: str, end: str, increase: bool):
        """缩进选中区域 - 每行增加/减少相同缩进量"""
        start_line = int(start.split('.')[0])
        end_line = int(end.split('.')[0])
        indent_unit = self._get_indent_unit()
        
        # 代码块内使用4空格
        if self._is_in_code_block():
            indent_unit = '    '
        
        for line_num in range(start_line, end_line + 1):
            if increase:
                self._text.insert(f"{line_num}.0", indent_unit)
            else:
                line = self._text.get(f"{line_num}.0", f"{line_num}.end")
                indent = self._get_indent(line)
                if indent.startswith(indent_unit):
                    self._text.delete(f"{line_num}.0", f"{line_num}.{len(indent_unit)}")
                elif indent:
                    # 删除尽可能多的空白（最多indent_unit长度）
                    delete_len = min(len(indent), len(indent_unit))
                    self._text.delete(f"{line_num}.0", f"{line_num}.{delete_len}")
        
        return 'break'
    
    def _is_in_code_block(self) -> bool:
        """检查光标是否在代码块内"""
        try:
            cursor = self._text.index('insert')
            line_num = int(cursor.split('.')[0])
            
            # 向上搜索代码块开始标记
            code_block_start = None
            for i in range(line_num, 0, -1):
                line = self._text.get(f"{i}.0", f"{i}.end")
                if line.strip().startswith('```'):
                    code_block_start = i
                    break
            
            if code_block_start is None:
                return False
            
            # 检查是否有对应的结束标记
            total_lines = int(self._text.index('end').split('.')[0])
            for i in range(code_block_start + 1, min(line_num + 1, total_lines)):
                line = self._text.get(f"{i}.0", f"{i}.end")
                if line.strip() == '```':
                    return False  # 代码块已结束
            
            return True
        except:
            return False

    def _on_backspace(self, event=None):
        """退格键处理"""
        if not self._enabled:
            return
        
        cursor = self._text.index('insert')
        col = int(cursor.split('.')[1])
        
        if col == 0:
            return  # 行首，使用默认行为
        
        line, line_num = self._get_current_line()
        indent = self._get_indent(line)
        
        # 如果光标在缩进区域内，删除整个缩进单位
        if col <= len(indent) and col > 0:
            indent_unit = self._get_indent_unit()
            if col >= len(indent_unit):
                # 删除一个缩进单位
                self._text.delete(f"{line_num}.{col - len(indent_unit)}", cursor)
                return 'break'
    
    def set_indent_size(self, size: int):
        """设置缩进大小"""
        self._indent_size = size
    
    def set_use_tabs(self, use_tabs: bool):
        """设置是否使用 Tab"""
        self._use_tabs = use_tabs
    
    def enable(self):
        """启用智能缩进"""
        self._enabled = True
    
    def disable(self):
        """禁用智能缩进"""
        self._enabled = False


class BracketMatcher:
    """括号匹配和自动补全 - 支持中英文括号"""
    
    def __init__(self, text_widget):
        """
        初始化括号匹配器
        
        Args:
            text_widget: tkinter Text 或 CTkTextbox 组件
        """
        self.text_widget = text_widget
        self._enabled = True
        self._auto_close = True
        self._highlight_pairs = True
        
        # 获取底层 Text 组件
        if hasattr(text_widget, '_textbox'):
            self._text = text_widget._textbox
        else:
            self._text = text_widget
        
        # 配置高亮标签
        self._text.tag_configure('bracket_match', background='#fef3c7')
        self._text.tag_configure('bracket_mismatch', background='#fee2e2')
        
        # 绑定事件
        self._bind_events()

    def _bind_events(self):
        """绑定键盘事件"""
        # 使用通用的按键事件处理中英文括号
        self._text.bind('<KeyPress>', self._on_key_press, add='+')
        
        # 退格删除配对括号
        self._text.bind('<BackSpace>', self._on_backspace, add='+')
        
        # 光标移动时更新高亮
        self._text.bind('<KeyRelease>', self._update_highlight)
        self._text.bind('<ButtonRelease-1>', self._update_highlight)
    
    def _on_key_press(self, event):
        """通用按键处理 - 处理所有括号输入"""
        if not self._enabled or not self._auto_close:
            return
        
        char = event.char
        if not char:
            return
        
        # 检查是否是左括号
        if char in BRACKET_PAIRS:
            return self._on_bracket_input(char)
        
        # 检查是否是右括号
        close_chars = set(BRACKET_PAIRS.values())
        if char in close_chars and char not in BRACKET_PAIRS:
            return self._on_close_bracket(char)
    
    def _on_bracket_input(self, char: str):
        """括号输入处理"""
        if not self._enabled or not self._auto_close:
            return
        
        # 检查是否有选中内容
        try:
            sel_start = self._text.index('sel.first')
            sel_end = self._text.index('sel.last')
            selected = self._text.get(sel_start, sel_end)
            
            # 用括号包裹选中内容
            close_char = BRACKET_PAIRS[char]
            self._text.delete(sel_start, sel_end)
            self._text.insert(sel_start, char + selected + close_char)
            
            # 选中包裹后的内容
            self._text.tag_add('sel', f"{sel_start}+1c", f"{sel_start}+{len(selected) + 1}c")
            
            return 'break'
        except tk.TclError:
            pass
        
        # 自动补全右括号
        close_char = BRACKET_PAIRS[char]
        
        # 对于引号类字符，检查是否应该关闭
        quote_chars = {'"', "'", '`', '\u201c', '\u2018'}  # 包含中文引号
        if char in quote_chars:
            cursor = self._text.index('insert')
            # 检查前一个字符
            try:
                prev_char = self._text.get(f"{cursor}-1c", cursor)
                if prev_char.isalnum():
                    return  # 不自动补全
            except:
                pass
        
        self._text.insert('insert', char + close_char)
        self._text.mark_set('insert', f'insert-1c')
        return 'break'
    
    def _on_close_bracket(self, char: str):
        """右括号输入处理"""
        if not self._enabled:
            return
        
        cursor = self._text.index('insert')
        next_char = self._text.get(cursor, f"{cursor}+1c")
        
        # 如果下一个字符就是要输入的右括号，跳过而不是插入
        if next_char == char:
            self._text.mark_set('insert', f'{cursor}+1c')
            return 'break'

    def _on_backspace(self, event=None):
        """退格键处理 - 删除配对括号"""
        if not self._enabled:
            return
        
        cursor = self._text.index('insert')
        
        try:
            prev_char = self._text.get(f"{cursor}-1c", cursor)
            next_char = self._text.get(cursor, f"{cursor}+1c")
            
            # 检查是否是配对括号
            if prev_char in BRACKET_PAIRS and BRACKET_PAIRS[prev_char] == next_char:
                self._text.delete(f"{cursor}-1c", f"{cursor}+1c")
                return 'break'
        except:
            pass
    
    def _update_highlight(self, event=None):
        """更新括号高亮"""
        if not self._enabled or not self._highlight_pairs:
            return
        
        # 清除现有高亮
        self._text.tag_remove('bracket_match', '1.0', 'end')
        self._text.tag_remove('bracket_mismatch', '1.0', 'end')
        
        cursor = self._text.index('insert')
        
        # 检查光标前后的字符
        for offset in ['-1c', '']:
            try:
                pos = f"{cursor}{offset}" if offset else cursor
                char = self._text.get(pos, f"{pos}+1c")
                
                if char in BRACKET_PAIRS:
                    # 找配对的右括号
                    match_pos = self._find_matching_bracket(pos, char, BRACKET_PAIRS[char], forward=True)
                    if match_pos:
                        self._text.tag_add('bracket_match', pos, f"{pos}+1c")
                        self._text.tag_add('bracket_match', match_pos, f"{match_pos}+1c")
                    else:
                        self._text.tag_add('bracket_mismatch', pos, f"{pos}+1c")
                    break
                
                elif char in BRACKET_PAIRS.values():
                    # 找配对的左括号
                    open_char = None
                    for k, v in BRACKET_PAIRS.items():
                        if v == char:
                            open_char = k
                            break
                    if open_char:
                        match_pos = self._find_matching_bracket(pos, char, open_char, forward=False)
                        if match_pos:
                            self._text.tag_add('bracket_match', pos, f"{pos}+1c")
                            self._text.tag_add('bracket_match', match_pos, f"{match_pos}+1c")
                        else:
                            self._text.tag_add('bracket_mismatch', pos, f"{pos}+1c")
                    break
            except:
                pass

    def _find_matching_bracket(self, start_pos: str, start_char: str, 
                                target_char: str, forward: bool) -> Optional[str]:
        """查找配对括号"""
        content = self._text.get('1.0', 'end')
        start_index = self._pos_to_index(start_pos, content)
        
        if start_index is None:
            return None
        
        count = 1
        i = start_index + (1 if forward else -1)
        
        while 0 <= i < len(content):
            char = content[i]
            
            if char == start_char:
                count += 1
            elif char == target_char:
                count -= 1
                if count == 0:
                    return self._index_to_pos(i, content)
            
            i += 1 if forward else -1
        
        return None
    
    def _pos_to_index(self, pos: str, content: str) -> Optional[int]:
        """将 tkinter 位置转换为字符串索引"""
        try:
            line, col = map(int, pos.split('.'))
            lines = content.split('\n')
            index = sum(len(lines[i]) + 1 for i in range(line - 1)) + col
            return index
        except:
            return None
    
    def _index_to_pos(self, index: int, content: str) -> str:
        """将字符串索引转换为 tkinter 位置"""
        lines = content.split('\n')
        current_index = 0
        
        for line_num, line in enumerate(lines, 1):
            if current_index + len(line) >= index:
                col = index - current_index
                return f"{line_num}.{col}"
            current_index += len(line) + 1
        
        return "1.0"
    
    def set_auto_close(self, enabled: bool):
        """设置是否自动补全括号"""
        self._auto_close = enabled
    
    def set_highlight_pairs(self, enabled: bool):
        """设置是否高亮配对括号"""
        self._highlight_pairs = enabled
        if not enabled:
            self._text.tag_remove('bracket_match', '1.0', 'end')
            self._text.tag_remove('bracket_mismatch', '1.0', 'end')
    
    def enable(self):
        """启用括号匹配"""
        self._enabled = True
    
    def disable(self):
        """禁用括号匹配"""
        self._enabled = False
        self._text.tag_remove('bracket_match', '1.0', 'end')
        self._text.tag_remove('bracket_mismatch', '1.0', 'end')


class CommentToggle:
    """注释切换功能 - Ctrl+/ 切换行注释"""
    
    def __init__(self, text_widget):
        """
        初始化注释切换
        
        Args:
            text_widget: tkinter Text 或 CTkTextbox 组件
        """
        self.text_widget = text_widget
        self._enabled = True
        
        # 获取底层 Text 组件
        if hasattr(text_widget, '_textbox'):
            self._text = text_widget._textbox
        else:
            self._text = text_widget
        
        # 绑定快捷键
        self._bind_events()
    
    def _bind_events(self):
        """绑定键盘事件"""
        self._text.bind('<Control-slash>', self._on_toggle_comment)
        self._text.bind('<Control-/>',  self._on_toggle_comment)
    
    def _on_toggle_comment(self, event=None):
        """切换注释"""
        if not self._enabled:
            return
        
        # 获取当前语言上下文
        lang = self._detect_language()
        comment_start, comment_end = COMMENT_FORMATS.get(lang, ('# ', ''))
        
        # 获取选中行或当前行
        try:
            sel_start = self._text.index('sel.first')
            sel_end = self._text.index('sel.last')
            start_line = int(sel_start.split('.')[0])
            end_line = int(sel_end.split('.')[0])
        except tk.TclError:
            cursor = self._text.index('insert')
            start_line = end_line = int(cursor.split('.')[0])
        
        # 检查是否所有行都已注释
        all_commented = True
        for line_num in range(start_line, end_line + 1):
            line = self._text.get(f"{line_num}.0", f"{line_num}.end")
            stripped = line.lstrip()
            if stripped and not stripped.startswith(comment_start.strip()):
                all_commented = False
                break
        
        # 切换注释
        for line_num in range(start_line, end_line + 1):
            line = self._text.get(f"{line_num}.0", f"{line_num}.end")
            if not line.strip():
                continue  # 跳过空行
            
            if all_commented:
                # 取消注释
                self._uncomment_line(line_num, comment_start, comment_end)
            else:
                # 添加注释
                self._comment_line(line_num, comment_start, comment_end)
        
        return 'break'

    def _detect_language(self) -> str:
        """检测当前语言上下文（基于代码块）"""
        try:
            cursor = self._text.index('insert')
            line_num = int(cursor.split('.')[0])
            
            # 向上搜索代码块开始标记
            for i in range(line_num, 0, -1):
                line = self._text.get(f"{i}.0", f"{i}.end")
                if line.strip().startswith('```'):
                    # 提取语言标识
                    lang = line.strip()[3:].strip().lower()
                    if lang:
                        return lang
                    break
            
            # 默认使用 Python 风格注释
            return 'python'
        except:
            return 'python'
    
    def _comment_line(self, line_num: int, comment_start: str, comment_end: str):
        """添加注释到指定行"""
        line = self._text.get(f"{line_num}.0", f"{line_num}.end")
        indent = len(line) - len(line.lstrip())
        
        if comment_end:
            # 块注释风格 (如 HTML, CSS)
            new_line = line[:indent] + comment_start + line[indent:] + comment_end
        else:
            # 行注释风格
            new_line = line[:indent] + comment_start + line[indent:]
        
        self._text.delete(f"{line_num}.0", f"{line_num}.end")
        self._text.insert(f"{line_num}.0", new_line)
    
    def _uncomment_line(self, line_num: int, comment_start: str, comment_end: str):
        """取消指定行的注释"""
        line = self._text.get(f"{line_num}.0", f"{line_num}.end")
        indent = len(line) - len(line.lstrip())
        content = line[indent:]
        
        # 移除注释标记
        if content.startswith(comment_start):
            content = content[len(comment_start):]
        
        if comment_end and content.endswith(comment_end):
            content = content[:-len(comment_end)]
        
        new_line = line[:indent] + content
        self._text.delete(f"{line_num}.0", f"{line_num}.end")
        self._text.insert(f"{line_num}.0", new_line)
    
    def enable(self):
        """启用注释切换"""
        self._enabled = True
    
    def disable(self):
        """禁用注释切换"""
        self._enabled = False


class SmartEditor:
    """智能编辑器 - 整合所有智能编辑功能"""
    
    def __init__(self, text_widget):
        """
        初始化智能编辑器
        
        Args:
            text_widget: tkinter Text 或 CTkTextbox 组件
        """
        self.text_widget = text_widget
        
        # 初始化各个功能模块
        self.smart_indent = SmartIndent(text_widget)
        self.bracket_matcher = BracketMatcher(text_widget)
        self.comment_toggle = CommentToggle(text_widget)
    
    def enable_all(self):
        """启用所有功能"""
        self.smart_indent.enable()
        self.bracket_matcher.enable()
        self.comment_toggle.enable()
    
    def disable_all(self):
        """禁用所有功能"""
        self.smart_indent.disable()
        self.bracket_matcher.disable()
        self.comment_toggle.disable()
    
    def configure(self, **kwargs):
        """配置选项"""
        if 'indent_size' in kwargs:
            self.smart_indent.set_indent_size(kwargs['indent_size'])
        if 'use_tabs' in kwargs:
            self.smart_indent.set_use_tabs(kwargs['use_tabs'])
        if 'auto_close_brackets' in kwargs:
            self.bracket_matcher.set_auto_close(kwargs['auto_close_brackets'])
        if 'highlight_brackets' in kwargs:
            self.bracket_matcher.set_highlight_pairs(kwargs['highlight_brackets'])
