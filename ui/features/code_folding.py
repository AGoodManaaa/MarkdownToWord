# -*- coding: utf-8 -*-
"""代码折叠模块 - 支持折叠代码块和标题区域"""

import re
import tkinter as tk
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass

try:
    import customtkinter as ctk
except ImportError:
    ctk = None


@dataclass
class FoldRegion:
    """可折叠区域"""
    start_line: int
    end_line: int
    region_type: str  # 'code_block', 'heading', 'list'
    is_folded: bool = False
    preview_text: str = ""  # 折叠时显示的预览文本


class CodeFolding:
    """代码折叠功能"""
    
    def __init__(self, text_widget, line_numbers_widget=None):
        """
        初始化代码折叠
        
        Args:
            text_widget: tkinter Text 或 CTkTextbox 组件
            line_numbers_widget: 行号组件 (Text 或 Canvas)
        """
        self.text_widget = text_widget
        self._enabled = True
        self._fold_regions: Dict[int, FoldRegion] = {}  # start_line -> FoldRegion
        self._folded_content: Dict[int, str] = {}  # start_line -> 原始内容
        
        # 获取底层 Text 组件
        if hasattr(text_widget, '_textbox'):
            self._text = text_widget._textbox
        else:
            self._text = text_widget
        
        # 行号组件
        self.line_widget = line_numbers_widget
        
        # 折叠图标
        self._fold_icon = "▼"
        self._unfold_icon = "▶"
        
        # 配置折叠标签样式
        self._text.tag_configure('folded', 
            background='#f3f4f6',
            foreground='#6b7280',
            font=('Consolas', 10, 'italic')
        )
        self._text.tag_configure('fold_marker',
            foreground='#9ca3af',
            font=('Consolas', 10)
        )
        
        # 绑定事件
        self._bind_events()
    
    def _bind_events(self):
        """绑定事件"""
        self._text.bind('<KeyRelease>', self._on_text_change)
        # 移除 <<Modified>> 绑定，改由 gui.py 统一分发驱动，避免竞争
        
        # 双击折叠区域展开
        self._text.tag_bind('folded', '<Double-Button-1>', self._on_fold_double_click)
    
    def _on_modified(self, event=None):
        """文本修改事件"""
        if self._text.edit_modified():
            self._text.edit_modified(False)
            self._scan_fold_regions()
    
    def _on_text_change(self, event=None):
        """文本变化事件"""
        if not self._enabled:
            return
        self._scan_fold_regions()
    
    def _scan_fold_regions(self):
        """扫描可折叠区域"""
        if not self._enabled:
            return
        
        content = self._text.get("1.0", "end-1c")
        lines = content.split('\n')
        
        new_regions = {}
        
        # 扫描代码块
        in_code_block = False
        code_block_start = 0
        code_block_lang = ""
        
        for i, line in enumerate(lines, 1):
            # 代码块开始
            if re.match(r'^```(\w*)$', line.strip()):
                if not in_code_block:
                    in_code_block = True
                    code_block_start = i
                    match = re.match(r'^```(\w*)$', line.strip())
                    code_block_lang = match.group(1) if match else ""
                else:
                    # 代码块结束
                    in_code_block = False
                    if i > code_block_start + 1:  # 至少有一行内容
                        line_count = i - code_block_start - 1
                        preview = f"... {line_count} 行代码"
                        if code_block_lang:
                            preview = f"[{code_block_lang}] {preview}"
                        
                        # 保留已有的折叠状态
                        old_region = self._fold_regions.get(code_block_start)
                        is_folded = old_region.is_folded if old_region else False
                        
                        new_regions[code_block_start] = FoldRegion(
                            start_line=code_block_start,
                            end_line=i,
                            region_type='code_block',
                            is_folded=is_folded,
                            preview_text=preview
                        )
        
        self._fold_regions = new_regions
        self._update_fold_markers()
    
    def _update_fold_markers(self):
        """更新折叠标记显示"""
        if not self.line_widget:
            return
        
        # 情况1: 如果行号栏是 Canvas
        if isinstance(self.line_widget, tk.Canvas):
            self.line_widget.delete("fold_icon")
            for start_line, region in self._fold_regions.items():
                try:
                    bbox = self._text.bbox(f"{start_line}.0")
                    if bbox:
                        y = bbox[1] + bbox[3] // 2
                        icon = self._unfold_icon if region.is_folded else self._fold_icon
                        self.line_widget.create_text(
                            10, y, text=icon, anchor='w', fill='#6b7280',
                            font=('Consolas', 8), tags=("fold_icon", f"fold_{start_line}")
                        )
                        self.line_widget.tag_bind(f"fold_{start_line}", '<Button-1>', 
                                                lambda e, sl=start_line: self.toggle_fold(sl))
                except: pass
        
        # 情况2: 如果行号栏是 Text (LineNumberedText 风格)
        elif isinstance(self.line_widget, tk.Text):
            # Text 风格下，图标通常显示在行号旁边
            # 这种风格由于 LineNumberedText 会全量刷新行号，折叠图标会被覆盖
            # 建议在 LineNumberedText 的 _update_line_numbers 中集成此逻辑
            pass
    
    def toggle_fold(self, start_line: int):
        """切换折叠状态"""
        region = self._fold_regions.get(start_line)
        if not region:
            return
        
        if region.is_folded:
            self._unfold(start_line)
        else:
            self._fold(start_line)
    
    def _fold(self, start_line: int):
        """折叠区域"""
        region = self._fold_regions.get(start_line)
        if not region or region.is_folded:
            return
        
        try:
            # 保存原始内容
            content_start = f"{start_line + 1}.0"
            content_end = f"{region.end_line}.0"
            original_content = self._text.get(content_start, content_end)
            self._folded_content[start_line] = original_content
            
            # 删除内容并插入折叠提示
            self._text.delete(content_start, content_end)
            self._text.insert(content_start, f"{region.preview_text}\n", ('folded',))
            
            # 更新状态
            region.is_folded = True
            self._update_fold_markers()
            
        except Exception as e:
            print(f"折叠失败: {e}")
    
    def _unfold(self, start_line: int):
        """展开区域"""
        region = self._fold_regions.get(start_line)
        if not region or not region.is_folded:
            return
        
        original_content = self._folded_content.get(start_line)
        if not original_content:
            return
        
        try:
            # 删除折叠提示
            fold_line = f"{start_line + 1}.0"
            fold_end = f"{start_line + 2}.0"
            self._text.delete(fold_line, fold_end)
            
            # 恢复原始内容
            self._text.insert(fold_line, original_content)
            
            # 更新状态
            region.is_folded = False
            del self._folded_content[start_line]
            self._update_fold_markers()
            
        except Exception as e:
            print(f"展开失败: {e}")
    
    def _on_fold_double_click(self, event):
        """双击折叠区域展开"""
        # 获取点击位置的行号
        index = self._text.index(f"@{event.x},{event.y}")
        line = int(index.split('.')[0])
        
        # 查找对应的折叠区域
        for start_line, region in self._fold_regions.items():
            if region.is_folded and line == start_line + 1:
                self._unfold(start_line)
                return 'break'
    
    def fold_all(self):
        """折叠所有区域"""
        for start_line in list(self._fold_regions.keys()):
            if not self._fold_regions[start_line].is_folded:
                self._fold(start_line)
    
    def unfold_all(self):
        """展开所有区域"""
        for start_line in list(self._fold_regions.keys()):
            if self._fold_regions[start_line].is_folded:
                self._unfold(start_line)
    
    def get_fold_regions(self) -> List[FoldRegion]:
        """获取所有可折叠区域"""
        return list(self._fold_regions.values())
    
    def enable(self):
        """启用代码折叠"""
        self._enabled = True
        self._scan_fold_regions()
    
    def disable(self):
        """禁用代码折叠"""
        self._enabled = False
        self.unfold_all()
        if self.line_canvas:
            self.line_canvas.delete("fold_icon")
