# -*- coding: utf-8 -*-
"""
迷你地图模块 - VS Code 风格文档缩略图导航
真实缩小版：内容按比例缩放显示，视口精确跟踪
"""

import tkinter as tk
from typing import Optional


class Minimap:
    """
    迷你地图 - 文档的真实缩小版
    
    特点：
    1. 内容按比例缩放，真实反映文档结构
    2. 视口指示器精确反映可见区域
    3. 无残影的平滑滚动
    4. VS Code 风格的交互体验
    """
    
    def __init__(self, text_widget, parent=None):
        self.text_widget = text_widget
        self._enabled = True
        self._visible = False
        self._width = 100
        self._update_delay = 30
        self._update_id = None
        self._viewport_update_id = None
        
        if hasattr(text_widget, '_textbox'):
            self._text = text_widget._textbox
        else:
            self._text = text_widget
        
        self._parent = parent or self._text.master
        
        # VS Code 风格颜色
        self._bg_color = '#f5f5f5'
        self._viewport_fill = '#d0d0d0'
        self._viewport_border = '#a0a0a0'
        self._viewport_hover = '#c0c0c0'
        
        self._container = tk.Frame(
            self._parent,
            width=self._width,
            bg=self._bg_color,
            highlightthickness=0,
        )
        
        self.canvas = tk.Canvas(
            self._container,
            width=self._width,
            bg=self._bg_color,
            highlightthickness=0,
            cursor='hand2',
            bd=0,
        )
        self.canvas.pack(fill='both', expand=True)

        self._separator = tk.Frame(
            self._container,
            width=1,
            bg='#e0e0e0',
        )
        
        self._is_hovering = False
        self._is_dragging = False
        self._drag_offset = 0
        
        self._total_lines = 0
        self._canvas_height = 0
        self._line_height = 2
        self._content_height = 0
        self._last_yview = (0, 1)
        
        self._bind_events()
    
    def _bind_events(self):
        """绑定事件"""
        self._text.bind('<KeyRelease>', self._schedule_update, add='+')
        self._text.bind('<<Modified>>', self._on_modified, add='+')
        self._text.bind('<Configure>', self._schedule_update, add='+')
        self._text.bind('<MouseWheel>', self._on_scroll, add='+')
        self._text.bind('<Button-4>', self._on_scroll, add='+')
        self._text.bind('<Button-5>', self._on_scroll, add='+')
        
        self.canvas.bind('<Button-1>', self._on_click)
        self.canvas.bind('<B1-Motion>', self._on_drag)
        self.canvas.bind('<ButtonRelease-1>', self._on_release)
        self.canvas.bind('<Enter>', self._on_enter)
        self.canvas.bind('<Leave>', self._on_leave)
        self.canvas.bind('<MouseWheel>', self._on_minimap_scroll)
        self.canvas.bind('<Configure>', self._on_canvas_configure)
    
    def _on_modified(self, event=None):
        try:
            if self._text.edit_modified():
                self._text.edit_modified(False)
                self._schedule_update()
        except:
            pass
    
    def _on_scroll(self, event=None):
        self._schedule_viewport_update()
    
    def _on_minimap_scroll(self, event):
        if event.delta > 0:
            self._text.yview_scroll(-3, 'units')
        else:
            self._text.yview_scroll(3, 'units')
        self._schedule_viewport_update()
        return 'break'
    
    def _on_canvas_configure(self, event=None):
        self._schedule_update()
    
    def _schedule_update(self, event=None):
        if self._update_id:
            self.canvas.after_cancel(self._update_id)
        self._update_id = self.canvas.after(self._update_delay, self._full_update)
    
    def _schedule_viewport_update(self):
        if self._viewport_update_id:
            self.canvas.after_cancel(self._viewport_update_id)
        self._viewport_update_id = self.canvas.after(10, self._update_viewport)
    
    def _full_update(self):
        self._update_id = None
        if not self._enabled or not self._visible:
            return
        
        self.canvas.delete("all")
        
        try:
            content = self._text.get("1.0", "end-1c")
        except:
            return
        
        lines = content.split('\n')
        self._total_lines = len(lines)
        
        if self._total_lines == 0:
            return
        
        self._canvas_height = self.canvas.winfo_height()
        if self._canvas_height <= 1:
            self._canvas_height = 400
        
        self._content_height = self._total_lines * self._line_height
        self._draw_content(lines)
        self._draw_viewport()
    
    def _draw_content(self, lines):
        y = 0
        for line in lines:
            color, width = self._get_line_style(line)
            if width > 0:
                self.canvas.create_rectangle(
                    4, y,
                    4 + min(width, self._width - 8), y + self._line_height - 1,
                    fill=color,
                    outline='',
                    tags='content'
                )
            y += self._line_height
    
    def _update_viewport(self):
        self._viewport_update_id = None
        if not self._enabled or not self._visible:
            return
        self.canvas.delete('viewport')
        self._draw_viewport()

    def _draw_viewport(self):
        if self._total_lines == 0 or self._content_height == 0:
            return
        
        try:
            yview = self._text.yview()
            self._last_yview = yview
            
            top_ratio = yview[0]
            bottom_ratio = yview[1]
            
            y1 = top_ratio * self._content_height
            y2 = bottom_ratio * self._content_height
            
            min_height = 20
            if y2 - y1 < min_height:
                center = (y1 + y2) / 2
                y1 = center - min_height / 2
                y2 = center + min_height / 2
            
            y1 = max(0, y1)
            y2 = min(self._content_height, y2)
            
            if self._is_dragging:
                fill = self._viewport_hover
                border = '#808080'
            elif self._is_hovering:
                fill = self._viewport_hover
                border = self._viewport_border
            else:
                fill = self._viewport_fill
                border = self._viewport_border
            
            self.canvas.create_rectangle(
                2, y1,
                self._width - 2, y2,
                fill=fill,
                outline=border,
                width=1,
                tags='viewport'
            )
            
            self.canvas.tag_lower('viewport')
        except Exception:
            pass
    
    def _get_line_style(self, line: str) -> tuple:
        stripped = line.strip()
        
        if not stripped:
            return '#e5e7eb', 0
        
        base_width = min(len(line) * 0.6, self._width - 12)
        
        if stripped.startswith('#'):
            level = len(stripped) - len(stripped.lstrip('#'))
            width = max(25, self._width - 12 - (level - 1) * 8)
            return '#4a90d9', width
        
        if stripped.startswith('```'):
            return '#2d8a56', self._width - 12
        
        if stripped.startswith(('-', '*', '+')) or (len(stripped) > 0 and stripped[0].isdigit() and '.' in stripped[:3]):
            return '#d97706', base_width
        
        if stripped.startswith('>'):
            return '#6b7280', base_width * 0.9
        
        if stripped.startswith('|'):
            return '#7c3aed', self._width - 12
        
        if '[' in stripped and ']' in stripped:
            return '#0891b2', base_width
        
        return '#9ca3af', base_width
    
    def _on_click(self, event):
        self._is_dragging = True
        
        if self._content_height == 0:
            return
        
        yview = self._text.yview()
        viewport_height_ratio = yview[1] - yview[0]
        viewport_height_px = viewport_height_ratio * self._content_height
        
        current_top = yview[0] * self._content_height
        current_bottom = yview[1] * self._content_height
        
        if current_top <= event.y <= current_bottom:
            self._drag_offset = event.y - current_top
        else:
            self._drag_offset = viewport_height_px / 2
            target_top = (event.y - self._drag_offset) / self._content_height
            target_top = max(0, min(1 - viewport_height_ratio, target_top))
            self._text.yview_moveto(target_top)
        
        self._update_viewport()
    
    def _on_drag(self, event):
        if not self._is_dragging or self._content_height == 0:
            return
        
        yview = self._text.yview()
        viewport_height_ratio = yview[1] - yview[0]
        
        new_top = (event.y - self._drag_offset) / self._content_height
        new_top = max(0, min(1 - viewport_height_ratio, new_top))
        
        self._text.yview_moveto(new_top)
        self._update_viewport()
    
    def _on_release(self, event):
        self._is_dragging = False
        self._update_viewport()
    
    def _on_enter(self, event):
        self._is_hovering = True
        self._update_viewport()
    
    def _on_leave(self, event):
        self._is_hovering = False
        self._is_dragging = False
        self._update_viewport()

    def show(self):
        if self._visible:
            return
        
        self._visible = True
        self._enabled = True
        
        self._container.place(
            relx=1.0,
            rely=0,
            relheight=1.0,
            width=self._width,
            anchor='ne'
        )
        
        self._separator.place(
            relx=1.0,
            rely=0,
            relheight=1.0,
            width=1,
            x=-self._width,
            anchor='ne'
        )
        
        self.canvas.after(50, self._full_update)
    
    def hide(self):
        self._visible = False
        self._enabled = False
        self._container.place_forget()
        self._separator.place_forget()
    
    def toggle(self):
        if self._visible:
            self.hide()
        else:
            self.show()
    
    def is_visible(self) -> bool:
        return self._visible
    
    def set_width(self, width: int):
        self._width = width
        self._container.configure(width=width)
        self.canvas.configure(width=width)
        if self._visible:
            self._container.place_configure(width=width)
            self._separator.place_configure(x=-width)
        self._full_update()
    
    def refresh(self):
        self._full_update()
