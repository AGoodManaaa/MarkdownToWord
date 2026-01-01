# -*- coding: utf-8 -*-
"""虚拟化渲染模块 - 只渲染可视区域的语法高亮"""

import re
import time
import tkinter as tk
from typing import Dict, List, Optional, Tuple, Set
from dataclasses import dataclass, field


@dataclass
class RenderBlock:
    """渲染块信息"""
    start_line: int
    end_line: int
    content_hash: str
    rendered: bool = False
    tags: List[Tuple[str, str, str]] = field(default_factory=list)  # [(tag, start, end), ...]


class VirtualRenderer:
    """虚拟化语法高亮渲染器 - 只渲染可视区域"""
    
    def __init__(self, text_widget, buffer_lines: int = 20):
        """
        初始化虚拟化渲染器
        
        Args:
            text_widget: tkinter Text 组件
            buffer_lines: 可视区域外的缓冲行数
        """
        self._text = text_widget
        self._buffer_lines = buffer_lines
        
        # 渲染缓存
        self._block_cache: Dict[int, RenderBlock] = {}
        self._rendered_range: Tuple[int, int] = (0, 0)
        
        # 节流控制
        self._last_render_time = 0.0
        self._render_cooldown = 0.03  # 30ms
        self._pending_render_id = None
        
        # 内容哈希（用于检测变化）
        self._content_hash = ""
        
        # 语法模式
        self._patterns = self._compile_patterns()
        
        # 标签配置
        self._configure_tags()
        
        # 绑定滚动事件
        self._bind_events()
    
    def _compile_patterns(self) -> Dict[str, re.Pattern]:
        """编译语法正则表达式"""
        return {
            'heading': re.compile(r'^(#{1,6})\s+(.+)$'),
            'bold': re.compile(r'(\*\*|__)(.+?)\1'),
            'italic': re.compile(r'(?<!\*)(\*|_)(?!\*)(.+?)\1(?!\*)'),
            'code_inline': re.compile(r'`([^`\n]+)`'),
            'code_block_start': re.compile(r'^```(\w*)$'),
            'code_block_end': re.compile(r'^```$'),
            'link': re.compile(r'\[([^\]]+)\]\(([^)]+)\)'),
            'image': re.compile(r'!\[([^\]]*)\]\(([^)]+)\)'),
            'list_marker': re.compile(r'^(\s*)([-*+]|\d+\.)\s+'),
            'blockquote': re.compile(r'^(>\s*)+'),
            'hr': re.compile(r'^([-*_]){3,}\s*$'),
            'strikethrough': re.compile(r'~~(.+?)~~'),
            'math_inline': re.compile(r'\$([^$\n]+)\$'),
            'task_list': re.compile(r'^(\s*)([-*+])\s+\[([ xX])\]\s+'),
        }
    
    def _configure_tags(self):
        """配置文本标签样式"""
        t = self._text
        
        # 标题样式
        for i in range(1, 7):
            size = 24 - (i - 1) * 2
            t.tag_configure(f'vh{i}', foreground='#3b82f6', 
                           font=('Microsoft YaHei', size, 'bold'))
            t.tag_configure(f'vh{i}_marker', foreground='#9ca3af')
        
        # 基础样式
        t.tag_configure('vbold', font=('Microsoft YaHei', 11, 'bold'))
        t.tag_configure('vitalic', font=('Microsoft YaHei', 11, 'italic'), 
                       foreground='#6b7280')
        t.tag_configure('vcode', foreground='#dc2626', background='#f3f4f6',
                       font=('Consolas', 10))
        t.tag_configure('vcode_block', foreground='#059669', background='#f3f4f6',
                       font=('Consolas', 10))
        t.tag_configure('vlink', foreground='#2563eb', underline=True)
        t.tag_configure('vimage', foreground='#7c3aed')
        t.tag_configure('vlist_marker', foreground='#f59e0b', 
                       font=('Microsoft YaHei', 11, 'bold'))
        t.tag_configure('vblockquote', foreground='#6b7280', background='#f9fafb')
        t.tag_configure('vhr', foreground='#d1d5db')
        t.tag_configure('vstrikethrough', overstrike=True, foreground='#9ca3af')
        t.tag_configure('vmath', foreground='#7c3aed', font=('Cambria Math', 11))
        t.tag_configure('vtask_done', overstrike=True, foreground='#9ca3af')
        t.tag_configure('vtask_checkbox', foreground='#f59e0b')
    
    def _bind_events(self):
        """绑定事件"""
        self._text.bind('<Configure>', self._on_configure, add='+')
        self._text.bind('<MouseWheel>', self._on_scroll, add='+')
        self._text.bind('<Button-4>', self._on_scroll, add='+')
        self._text.bind('<Button-5>', self._on_scroll, add='+')
    
    def _on_configure(self, event=None):
        """窗口大小变化时重新渲染"""
        self._schedule_render()
    
    def _on_scroll(self, event=None):
        """滚动时按需渲染"""
        self._schedule_render()
    
    def _schedule_render(self):
        """调度渲染（带节流）"""
        now = time.monotonic()
        if now - self._last_render_time < self._render_cooldown:
            # 节流：取消之前的调度，重新调度
            if self._pending_render_id:
                try:
                    self._text.after_cancel(self._pending_render_id)
                except Exception:
                    pass
            delay = int((self._render_cooldown - (now - self._last_render_time)) * 1000)
            self._pending_render_id = self._text.after(max(1, delay), self._render_visible)
            return
        
        self._render_visible()
    
    def _render_visible(self):
        """渲染可视区域"""
        self._pending_render_id = None
        self._last_render_time = time.monotonic()
        
        try:
            # 获取可视区域
            first_visible = self._text.index("@0,0")
            last_visible = self._text.index(f"@0,{self._text.winfo_height()}")
            
            first_line = int(first_visible.split('.')[0])
            last_line = int(last_visible.split('.')[0])
            
            # 添加缓冲区
            render_start = max(1, first_line - self._buffer_lines)
            render_end = last_line + self._buffer_lines
            
            # 检查是否需要重新渲染
            if (render_start, render_end) == self._rendered_range:
                return
            
            # 清除旧的高亮（只清除不在新范围内的）
            self._clear_outside_range(render_start, render_end)
            
            # 渲染新范围
            self._render_range(render_start, render_end)
            
            self._rendered_range = (render_start, render_end)
            
        except Exception:
            pass
    
    def _clear_outside_range(self, start: int, end: int):
        """清除范围外的高亮"""
        old_start, old_end = self._rendered_range
        
        # 清除上方超出的部分
        if old_start < start:
            self._clear_range(old_start, start - 1)
        
        # 清除下方超出的部分
        if old_end > end:
            self._clear_range(end + 1, old_end)
    
    def _clear_range(self, start: int, end: int):
        """清除指定范围的高亮"""
        for tag in self._text.tag_names():
            if tag.startswith('v') and tag != 'sel':
                try:
                    self._text.tag_remove(tag, f"{start}.0", f"{end}.end")
                except Exception:
                    pass
    
    def _render_range(self, start: int, end: int):
        """渲染指定范围"""
        try:
            content = self._text.get(f"{start}.0", f"{end}.end")
            lines = content.split('\n')
            
            in_code_block = False
            
            for i, line in enumerate(lines):
                line_num = start + i
                
                # 代码块处理
                if self._patterns['code_block_start'].match(line):
                    in_code_block = True
                    self._text.tag_add('vcode_block', f"{line_num}.0", f"{line_num}.end")
                    continue
                
                if self._patterns['code_block_end'].match(line) and in_code_block:
                    in_code_block = False
                    self._text.tag_add('vcode_block', f"{line_num}.0", f"{line_num}.end")
                    continue
                
                if in_code_block:
                    self._text.tag_add('vcode_block', f"{line_num}.0", f"{line_num}.end")
                    continue
                
                # 普通行高亮
                self._highlight_line(line, line_num)
                
        except Exception:
            pass
    
    def _highlight_line(self, line: str, line_num: int):
        """高亮单行"""
        line_start = f"{line_num}.0"
        
        # 标题
        match = self._patterns['heading'].match(line)
        if match:
            level = len(match.group(1))
            marker_end = f"{line_num}.{len(match.group(1))}"
            self._text.tag_add(f'vh{level}_marker', line_start, marker_end)
            self._text.tag_add(f'vh{level}', marker_end, f"{line_num}.end")
            return
        
        # 分隔线
        if self._patterns['hr'].match(line):
            self._text.tag_add('vhr', line_start, f"{line_num}.end")
            return
        
        # 引用
        match = self._patterns['blockquote'].match(line)
        if match:
            self._text.tag_add('vblockquote', line_start, f"{line_num}.end")
        
        # 任务列表
        match = self._patterns['task_list'].match(line)
        if match:
            indent_len = len(match.group(1))
            checkbox_end = f"{line_num}.{match.end()}"
            self._text.tag_add('vlist_marker', f"{line_num}.{indent_len}", 
                              f"{line_num}.{indent_len + 1}")
            self._text.tag_add('vtask_checkbox', f"{line_num}.{indent_len + 2}", checkbox_end)
            if match.group(3).lower() == 'x':
                self._text.tag_add('vtask_done', checkbox_end, f"{line_num}.end")
            return
        
        # 列表标记
        match = self._patterns['list_marker'].match(line)
        if match:
            indent_len = len(match.group(1))
            marker_len = len(match.group(2))
            self._text.tag_add('vlist_marker', f"{line_num}.{indent_len}", 
                              f"{line_num}.{indent_len + marker_len}")
        
        # 行内元素
        self._highlight_inline(line, line_num)
    
    def _highlight_inline(self, line: str, line_num: int):
        """高亮行内元素"""
        # 图片
        for match in self._patterns['image'].finditer(line):
            self._text.tag_add('vimage', f"{line_num}.{match.start()}", 
                              f"{line_num}.{match.end()}")
        
        # 链接
        for match in self._patterns['link'].finditer(line):
            if match.start() > 0 and line[match.start() - 1] == '!':
                continue
            self._text.tag_add('vlink', f"{line_num}.{match.start()}", 
                              f"{line_num}.{match.end()}")
        
        # 数学公式
        for match in self._patterns['math_inline'].finditer(line):
            self._text.tag_add('vmath', f"{line_num}.{match.start()}", 
                              f"{line_num}.{match.end()}")
        
        # 行内代码
        for match in self._patterns['code_inline'].finditer(line):
            self._text.tag_add('vcode', f"{line_num}.{match.start()}", 
                              f"{line_num}.{match.end()}")
        
        # 粗体
        for match in self._patterns['bold'].finditer(line):
            self._text.tag_add('vbold', f"{line_num}.{match.start()}", 
                              f"{line_num}.{match.end()}")
        
        # 斜体
        for match in self._patterns['italic'].finditer(line):
            if match.start() > 0 and line[match.start() - 1] == '*':
                continue
            if match.end() < len(line) and line[match.end()] == '*':
                continue
            self._text.tag_add('vitalic', f"{line_num}.{match.start()}", 
                              f"{line_num}.{match.end()}")
        
        # 删除线
        for match in self._patterns['strikethrough'].finditer(line):
            self._text.tag_add('vstrikethrough', f"{line_num}.{match.start()}", 
                              f"{line_num}.{match.end()}")
    
    def refresh(self):
        """强制刷新渲染"""
        self._rendered_range = (0, 0)
        self._block_cache.clear()
        self._render_visible()
    
    def clear(self):
        """清除所有高亮"""
        for tag in self._text.tag_names():
            if tag.startswith('v') and tag != 'sel':
                try:
                    self._text.tag_remove(tag, "1.0", "end")
                except Exception:
                    pass
        self._rendered_range = (0, 0)
        self._block_cache.clear()


class ProgressiveLoader:
    """渐进式文档加载器 - 分块加载大文档"""
    
    def __init__(self, text_widget, chunk_size: int = 1000, delay_ms: int = 10):
        """
        初始化渐进式加载器
        
        Args:
            text_widget: tkinter Text 组件
            chunk_size: 每次加载的字符数
            delay_ms: 加载间隔（毫秒）
        """
        self._text = text_widget
        self._chunk_size = chunk_size
        self._delay_ms = delay_ms
        
        self._loading = False
        self._content = ""
        self._position = 0
        self._on_progress = None
        self._on_complete = None
    
    def load(self, content: str, on_progress=None, on_complete=None):
        """
        开始加载内容
        
        Args:
            content: 要加载的内容
            on_progress: 进度回调 (loaded, total)
            on_complete: 完成回调
        """
        if self._loading:
            return
        
        self._content = content
        self._position = 0
        self._on_progress = on_progress
        self._on_complete = on_complete
        self._loading = True
        
        # 清空现有内容
        self._text.delete("1.0", "end")
        
        # 开始分块加载
        self._load_chunk()
    
    def _load_chunk(self):
        """加载一个块"""
        if not self._loading:
            return
        
        total = len(self._content)
        
        if self._position >= total:
            # 加载完成
            self._loading = False
            if self._on_complete:
                self._on_complete()
            return
        
        # 加载下一块
        end = min(self._position + self._chunk_size, total)
        chunk = self._content[self._position:end]
        
        self._text.insert("end", chunk)
        self._position = end
        
        # 进度回调
        if self._on_progress:
            self._on_progress(self._position, total)
        
        # 调度下一块
        self._text.after(self._delay_ms, self._load_chunk)
    
    def cancel(self):
        """取消加载"""
        self._loading = False
    
    @property
    def is_loading(self) -> bool:
        """是否正在加载"""
        return self._loading
    
    @property
    def progress(self) -> float:
        """加载进度 (0.0 - 1.0)"""
        if not self._content:
            return 1.0
        return self._position / len(self._content)
