# -*- coding: utf-8 -*-
"""
迷你地图模块 - VS Code 风格文档缩略图导航
真实缩小版：内容按比例缩放显示
"""

import tkinter as tk
from typing import Optional


class Minimap:
    """
    迷你地图 - 文档的真实缩小版
    
    特点：
    1. 内容按比例缩放，不是固定高度
    2. 视口指示器精确反映可见区域
    3. 无残影的平滑滚动
    """
    
    def __init__(self, text_widget, parent=None):
        self.text_widget = text_widget
        self._enabled = True
        self._visible = False
        self._width = 80
        self._update_delay = 50
        self._update_id = None
        
        # 缩放比例（文档高度 -> 迷你地图高度）
        self._scale = 0.15  # 15% 缩放
        
        # 获取底层 Text 组件
        if hasattr(text_widget, '_textbox'):
            self._text = text_widget._textbox
        else:
            self._text = text_widget
        
        self._parent = parent or self._text.master
        
        # 颜色配置 - VS Code 风格
        self._bg_color = '#f8fafc'
        self._bg_hover = '#f1f5f9'
        self._viewport_color = '#cce4ff'  # 浅蓝色
        self._viewport_border = '#80bfff'  # 蓝色边框
        self._viewport_hover = '#b3d9ff'
        
        # 创建容器
        self._container = tk.Frame(
            self._parent,
            width=self._width,
            bg=self._bg_color,
            highlightthickness=0,
        )
        
        # 创建画布
        self.canvas = tk.Canvas(
            self._container,
            width=self._width,
            bg=self._bg_color,
            highlightthickness=0,
            cursor='hand2',
            bd=0,
        )
        self.canvas.pack(fill='both', expand=True)
        
        # 分隔线
        self._separator = tk.Frame(
            self._container,
            width=1,
            bg='#e2e8f0',
        )
        
        # 状态
        self._is_hovering = False
        self._is_dragging = False
        self._drag_offset = 0  # 拖动时鼠标相对视口顶部的偏移
        
        # 缓存
        self._total_lines = 0
        self._canvas_height = 0
        self._content_height = 0  # 缩放后的内容总高度
        self._line_height = 2  # 每行在迷你地图中的高度
        
        self._bind_events()
    
    def _bind_events(self):
        """绑定事件"""
        self._text.bind('<KeyRelease>', self._schedule_update, add='+')
        self._text.bind('<<Modified>>', self._on_modified, add='+')
        self._text.bind('<MouseWheel>', self._on_scroll, add='+')
        self._text.bind('<Button-4>', self._on_scroll, add='+')
        self._text.bind('<Button-5>', self._on_scroll, add='+')
        self._text.bind('<Configure>', self._schedule_update, add='+')
        
        self.canvas.bind('<Button-1>', self._on_click)
        self.canvas.bind('<B1-Motion>', self._on_drag)
        self.canvas.bind('<ButtonRelease-1>', self._on_release)
        self.canvas.bind('<Enter>', self._on_enter)
        self.canvas.bind('<Leave>', self._on_leave)
        self.canvas.bind('<MouseWheel>', self._on_minimap_scroll)
    
    def _on_modified(self, event=None):
        try:
            if self._text.edit_modified():
                self._text.edit_modified(False)
                self._schedule_update()
        except:
            pass
    
    def _on_scroll(self, event=None):
        # 直接更新视口，不重绘内容
        self.canvas.after_idle(self._update_viewport_only)
    
    def _on_minimap_scroll(self, event):
        if event.delta > 0:
            self._text.yview_scroll(-3, 'units')
        else:
            self._text.yview_scroll(3, 'units')
        self._update_viewport_only()
        return 'break'
    
    def _schedule_update(self, event=None):
        if self._update_id:
            self.canvas.after_cancel(self._update_id)
        self._update_id = self.canvas.after(self._update_delay, self._update)
    
    def _update(self):
        """完整更新迷你地图（内容+视口）"""
        if not self._enabled or not self._visible:
            return
        
        # 清除所有内容
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
        
        # 计算缩放后的内容高度
        # 每行固定高度，内容总高度 = 行数 * 行高
        self._line_height = 2  # 每行2像素
        self._content_height = self._total_lines * self._line_height
        
        # 绘制内容
        y = 0
        for i, line in enumerate(lines):
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
        
        # 绘制视口
        self._draw_viewport()
    
    def _update_viewport_only(self):
        """只更新视口位置（不重绘内容）"""
        if not self._enabled or not self._visible:
            return
        
        # 删除旧视口
        self.canvas.delete('viewport')
        # 绘制新视口
        self._draw_viewport()
    
    def _draw_viewport(self):
        """绘制视口指示器"""
        if self._total_lines == 0 or self._content_height == 0:
            return
        
        try:
            yview = self._text.yview()
            top_ratio = yview[0]
            bottom_ratio = yview[1]
            
            # 视口在内容中的位置
            y1 = top_ratio * self._content_height
            y2 = bottom_ratio * self._content_height
            
            # 最小高度
            min_height = 15
            if y2 - y1 < min_height:
                center = (y1 + y2) / 2
                y1 = center - min_height / 2
                y2 = center + min_height / 2
            
            # 边界检查
            y1 = max(0, y1)
            y2 = min(self._content_height, y2)
            
            # 颜色
            fill = self._viewport_hover if (self._is_hovering or self._is_dragging) else self._viewport_color
            
            # 绘制视口
            self.canvas.create_rectangle(
                1, y1, 
                self._width - 1, y2,
                fill=fill,
                outline=self._viewport_border,
                width=1,
                tags='viewport'
            )
            
            # 视口在底层
            self.canvas.tag_lower('viewport')
            
        except Exception:
            pass
    
    def _get_line_style(self, line: str) -> tuple:
        """根据行内容获取样式"""
        stripped = line.strip()
        
        if not stripped:
            return '#e5e7eb', 0
        
        base_width = min(len(line) * 0.5, self._width - 12)
        
        if stripped.startswith('#'):
            level = len(stripped) - len(stripped.lstrip('#'))
            width = max(20, self._width - 12 - (level - 1) * 5)
            return '#3b82f6', width
        
        if stripped.startswith('```'):
            return '#059669', self._width - 12
        
        if stripped.startswith(('-', '*', '+')) or (len(stripped) > 0 and stripped[0].isdigit() and '.' in stripped[:3]):
            return '#f59e0b', base_width
        
        if stripped.startswith('>'):
            return '#6b7280', base_width * 0.9
        
        if stripped.startswith('|'):
            return '#8b5cf6', self._width - 12
        
        if '[' in stripped and ']' in stripped:
            return '#06b6d4', base_width
        
        return '#cbd5e1', base_width
    
    def _on_click(self, event):
        """点击 - 跳转到对应位置"""
        self._is_dragging = True
        
        if self._content_height == 0:
            return
        
        # 获取当前视口信息
        yview = self._text.yview()
        viewport_height_ratio = yview[1] - yview[0]
        viewport_height_px = viewport_height_ratio * self._content_height
        
        # 检查是否点击在视口内
        current_top = yview[0] * self._content_height
        current_bottom = yview[1] * self._content_height
        
        if current_top <= event.y <= current_bottom:
            # 点击在视口内，记录偏移量用于拖动
            self._drag_offset = event.y - current_top
        else:
            # 点击在视口外，跳转到该位置（视口中心对准点击位置）
            self._drag_offset = viewport_height_px / 2
            target_top = (event.y - self._drag_offset) / self._content_height
            target_top = max(0, min(1 - viewport_height_ratio, target_top))
            self._text.yview_moveto(target_top)
        
        self._update_viewport_only()
    
    def _on_drag(self, event):
        """拖动 - 视口跟随鼠标"""
        if not self._is_dragging or self._content_height == 0:
            return
        
        yview = self._text.yview()
        viewport_height_ratio = yview[1] - yview[0]
        
        # 计算新的视口顶部位置
        new_top = (event.y - self._drag_offset) / self._content_height
        new_top = max(0, min(1 - viewport_height_ratio, new_top))
        
        self._text.yview_moveto(new_top)
        self._update_viewport_only()
    
    def _on_release(self, event):
        """释放鼠标"""
        self._is_dragging = False
        self._update_viewport_only()
    
    def _on_enter(self, event):
        self._is_hovering = True
        self.canvas.configure(bg=self._bg_hover)
        self._update_viewport_only()
    
    def _on_leave(self, event):
        self._is_hovering = False
        self._is_dragging = False
        self.canvas.configure(bg=self._bg_color)
        self._update_viewport_only()
    
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
        
        self.canvas.after(50, self._update)
    
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
        self._update()
    
    def refresh(self):
        self._update()
