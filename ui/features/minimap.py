# -*- coding: utf-8 -*-
"""迷你地图模块 - VS Code 风格文档缩略图导航"""

import tkinter as tk
from typing import Optional

try:
    import customtkinter as ctk
except ImportError:
    ctk = None


class Minimap:
    """迷你地图 - VS Code 风格，嵌入编辑区右侧"""
    
    def __init__(self, text_widget, parent=None):
        """
        初始化迷你地图
        
        Args:
            text_widget: 关联的文本组件
            parent: 父容器 (应该是包含文本组件的 Frame)
        """
        self.text_widget = text_widget
        self._enabled = True
        self._visible = False
        self._width = 80  # 迷你地图宽度
        self._update_delay = 100  # 更新延迟 (ms)
        self._update_id = None
        
        # 获取底层 Text 组件
        if hasattr(text_widget, '_textbox'):
            self._text = text_widget._textbox
        else:
            self._text = text_widget
        
        # 获取父容器
        self._parent = parent or self._text.master
        
        # 创建迷你地图容器 Frame（放在编辑区内部右侧）
        self._container = tk.Frame(
            self._parent,
            width=self._width,
            bg='#f8fafc',
            highlightthickness=0,
        )
        
        # 创建迷你地图画布
        self.canvas = tk.Canvas(
            self._container,
            width=self._width,
            bg='#f8fafc',
            highlightthickness=0,
            cursor='hand2',
            bd=0,
        )
        self.canvas.pack(fill='both', expand=True)
        
        # 左侧分隔线
        self._separator = tk.Frame(
            self._container,
            width=1,
            bg='#e2e8f0',
        )
        
        # 可见区域指示器
        self._viewport_rect = None
        self._viewport_color = '#a7f3d0'  # 浅绿色（Tkinter不支持透明度）
        self._viewport_border = '#10b981'
        self._hover_color = '#6ee7b7'  # 悬停时更明显
        
        # 鼠标悬停状态
        self._is_hovering = False
        
        # 绑定事件
        self._bind_events()
    
    def _bind_events(self):
        """绑定事件"""
        # 文本变化
        self._text.bind('<KeyRelease>', self._schedule_update, add='+')
        self._text.bind('<<Modified>>', self._on_modified, add='+')
        
        # 滚动
        self._text.bind('<MouseWheel>', self._on_scroll, add='+')
        self._text.bind('<Button-4>', self._on_scroll, add='+')
        self._text.bind('<Button-5>', self._on_scroll, add='+')
        self._text.bind('<Configure>', self._schedule_update, add='+')
        
        # 迷你地图交互
        self.canvas.bind('<Button-1>', self._on_click)
        self.canvas.bind('<B1-Motion>', self._on_drag)
        self.canvas.bind('<Enter>', self._on_enter)
        self.canvas.bind('<Leave>', self._on_leave)
        self.canvas.bind('<MouseWheel>', self._on_minimap_scroll)
    
    def _on_modified(self, event=None):
        """文本修改事件"""
        try:
            if self._text.edit_modified():
                self._text.edit_modified(False)
                self._schedule_update()
        except:
            pass
    
    def _on_scroll(self, event=None):
        """滚动事件"""
        self._text.after(10, self._update_viewport)
    
    def _on_minimap_scroll(self, event):
        """迷你地图上的滚轮事件 - 传递给文本组件"""
        # 将滚轮事件传递给文本组件
        if event.delta > 0:
            self._text.yview_scroll(-3, 'units')
        else:
            self._text.yview_scroll(3, 'units')
        self._update_viewport()
        return 'break'
    
    def _schedule_update(self, event=None):
        """调度更新（防抖）"""
        if self._update_id:
            self.canvas.after_cancel(self._update_id)
        self._update_id = self.canvas.after(self._update_delay, self._update)
    
    def _update(self):
        """更新迷你地图"""
        if not self._enabled or not self._visible:
            return
        
        self.canvas.delete("all")
        
        # 获取文本内容
        try:
            content = self._text.get("1.0", "end-1c")
        except:
            return
            
        lines = content.split('\n')
        total_lines = len(lines)
        
        if total_lines == 0:
            return
        
        # 计算缩放
        canvas_height = self.canvas.winfo_height()
        if canvas_height <= 1:
            canvas_height = 400
        
        # 每行高度（最小1像素）
        line_height = max(1, min(3, canvas_height / max(total_lines, 1)))
        
        # 绘制文本缩略图
        y = 2
        for i, line in enumerate(lines):
            if y > canvas_height - 2:
                break
            
            # 根据行内容确定颜色和宽度
            color, width = self._get_line_style(line)
            
            if width > 0:
                # 绘制小矩形表示文本行
                self.canvas.create_rectangle(
                    4, y, 
                    4 + min(width, self._width - 8), y + max(1, line_height - 0.5),
                    fill=color, 
                    outline='',
                    tags='line'
                )
            
            y += line_height
        
        # 更新可见区域指示器
        self._update_viewport()
    
    def _get_line_style(self, line: str) -> tuple:
        """根据行内容获取样式 (颜色, 宽度)"""
        stripped = line.strip()
        
        if not stripped:
            return '#e5e7eb', 0  # 空行
        
        # 计算基础宽度（根据行长度）
        base_width = min(len(line) * 0.6, self._width - 12)
        
        # 标题 - 蓝色，较宽
        if stripped.startswith('#'):
            level = len(stripped) - len(stripped.lstrip('#'))
            width = max(30, self._width - 12 - (level - 1) * 8)
            return '#3b82f6', width
        
        # 代码块标记 - 绿色
        if stripped.startswith('```'):
            return '#059669', self._width - 12
        
        # 列表项 - 橙色
        if stripped.startswith(('-', '*', '+')) or (len(stripped) > 0 and stripped[0].isdigit() and '.' in stripped[:3]):
            return '#f59e0b', base_width
        
        # 引用 - 灰色，带缩进
        if stripped.startswith('>'):
            return '#6b7280', base_width * 0.9
        
        # 表格 - 紫色
        if stripped.startswith('|'):
            return '#8b5cf6', self._width - 12
        
        # 链接/图片 - 青色
        if '[' in stripped and ']' in stripped:
            return '#06b6d4', base_width
        
        # 普通文本 - 浅灰色
        return '#cbd5e1', base_width
    
    def _update_viewport(self):
        """更新可见区域指示器"""
        if not self._enabled or not self._visible:
            return
        
        # 删除旧的指示器
        self.canvas.delete('viewport')
        
        # 获取可见区域
        try:
            first_visible = self._text.index("@0,0")
            last_visible = self._text.index(f"@0,{self._text.winfo_height()}")
            
            first_line = int(first_visible.split('.')[0])
            last_line = int(last_visible.split('.')[0])
            total_lines = int(self._text.index('end-1c').split('.')[0])
            
            if total_lines == 0:
                return
            
            canvas_height = self.canvas.winfo_height()
            if canvas_height <= 1:
                canvas_height = 400
            
            # 计算指示器位置
            y1 = (first_line - 1) / total_lines * canvas_height
            y2 = last_line / total_lines * canvas_height
            
            # 确保最小高度
            min_height = 30
            if y2 - y1 < min_height:
                center = (y1 + y2) / 2
                y1 = center - min_height / 2
                y2 = center + min_height / 2
            
            # 边界检查
            y1 = max(0, y1)
            y2 = min(canvas_height, y2)
            
            # 选择颜色（悬停时更明显）
            fill_color = self._hover_color if self._is_hovering else self._viewport_color
            
            # 绘制指示器（圆角矩形效果）
            self._viewport_rect = self.canvas.create_rectangle(
                1, y1, 
                self._width - 1, y2,
                fill=fill_color,
                outline=self._viewport_border,
                width=1,
                tags='viewport'
            )
            
            # 将指示器放到底层，让文本行显示在上面
            self.canvas.tag_lower('viewport')
            
        except Exception as e:
            pass
    
    def _on_click(self, event):
        """点击事件 - 跳转到对应位置"""
        self._scroll_to_position(event.y)
    
    def _on_drag(self, event):
        """拖动事件 - 平滑滚动"""
        self._scroll_to_position(event.y)
    
    def _scroll_to_position(self, y: int):
        """滚动到指定位置"""
        canvas_height = self.canvas.winfo_height()
        if canvas_height <= 1:
            return
        
        try:
            total_lines = int(self._text.index('end-1c').split('.')[0])
            if total_lines == 0:
                return
            
            # 计算目标行（点击位置对应的行）
            ratio = y / canvas_height
            target_line = int(ratio * total_lines) + 1
            target_line = max(1, min(target_line, total_lines))
            
            # 滚动到目标行（居中显示）
            self._text.see(f"{target_line}.0")
            
            # 稍微延迟更新视口，确保滚动完成
            self._text.after(10, self._update_viewport)
        except:
            pass
    
    def _on_enter(self, event):
        """鼠标进入"""
        self._is_hovering = True
        self.canvas.configure(bg='#f1f5f9')
        self._update_viewport()
    
    def _on_leave(self, event):
        """鼠标离开"""
        self._is_hovering = False
        self.canvas.configure(bg='#f8fafc')
        self._update_viewport()
    
    def show(self):
        """显示迷你地图 - 放在编辑区内部右侧"""
        if self._visible:
            return
            
        self._visible = True
        self._enabled = True
        
        # 使用 place 将迷你地图放在父容器的右侧
        # 这样不会影响文本编辑区的布局
        self._container.place(
            relx=1.0,  # 右对齐
            rely=0,
            relheight=1.0,
            width=self._width,
            anchor='ne'  # 右上角锚点
        )
        
        # 显示分隔线
        self._separator.place(
            relx=1.0,
            rely=0,
            relheight=1.0,
            width=1,
            x=-self._width,
            anchor='ne'
        )
        
        # 延迟更新，确保布局完成
        self.canvas.after(50, self._update)
    
    def hide(self):
        """隐藏迷你地图"""
        self._visible = False
        self._enabled = False
        self._container.place_forget()
        self._separator.place_forget()
    
    def toggle(self):
        """切换显示状态"""
        if self._visible:
            self.hide()
        else:
            self.show()
    
    def is_visible(self) -> bool:
        """返回是否可见"""
        return self._visible
    
    def set_width(self, width: int):
        """设置宽度"""
        self._width = width
        self._container.configure(width=width)
        self.canvas.configure(width=width)
        if self._visible:
            self._container.place_configure(width=width)
            self._separator.place_configure(x=-width)
        self._update()
    
    def refresh(self):
        """强制刷新"""
        self._update()
