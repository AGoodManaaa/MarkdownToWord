# -*- coding: utf-8 -*-
"""远程光标管理模块 - 腾讯文档风格增强版

Features:
- 带用户名标签的彩色光标
- 光标闪烁动画
- 选区半透明高亮
- 光标入场/离场淡入淡出
- 不活跃时自动隐藏
"""

from dataclasses import dataclass, field
from typing import Dict, Optional, Tuple, List, Any
import time
import tkinter as tk

try:
    import customtkinter as ctk
    CTK_AVAILABLE = True
except ImportError:
    CTK_AVAILABLE = False


# 更丰富的协作光标配色 - 高对比度、现代感
CURSOR_COLORS = [
    {"primary": "#FF6B6B", "light": "#FFE5E5", "name": "珊瑚红"},
    {"primary": "#4ECDC4", "light": "#E0F7F5", "name": "薄荷绿"},
    {"primary": "#45B7D1", "light": "#E3F4F9", "name": "天空蓝"},
    {"primary": "#9B59B6", "light": "#F3E5F5", "name": "浪漫紫"},
    {"primary": "#F39C12", "light": "#FEF5E7", "name": "暖阳橙"},
    {"primary": "#E91E63", "light": "#FCE4EC", "name": "玫瑰粉"},
    {"primary": "#00BCD4", "light": "#E0F7FA", "name": "青碧"},
    {"primary": "#8BC34A", "light": "#F1F8E9", "name": "草木绿"},
    {"primary": "#FF5722", "light": "#FBE9E7", "name": "活力橙"},
    {"primary": "#3F51B5", "light": "#E8EAF6", "name": "靛蓝"},
]


@dataclass
class RemoteCursor:
    """远程光标 - 增强版"""
    participant_id: str
    name: str
    color: str
    color_light: str = "#E0E0E0"
    position: int = 0
    selection_start: Optional[int] = None
    selection_end: Optional[int] = None
    last_update: float = 0.0
    is_visible: bool = True
    opacity: float = 1.0
    
    @property
    def has_selection(self) -> bool:
        """是否有选区"""
        return self.selection_start is not None and self.selection_end is not None
    
    @property
    def is_active(self) -> bool:
        """是否活跃（5秒内有操作）"""
        return time.time() - self.last_update < 5.0


class CursorLabel:
    """光标用户名标签 - 腾讯文档风格
    
    显示在光标位置上方的彩色标签，包含用户名
    """
    
    def __init__(self, parent, name: str, color: str, x: int, y: int):
        self.parent = parent
        self.name = name
        self.color = color
        self.frame: Optional[Any] = None
        self.label: Optional[Any] = None
        self._x = x
        self._y = y
        self._opacity = 0.0
        self._animation_id = None
        
    def show(self, x: int = None, y: int = None):
        """显示标签"""
        if x is not None:
            self._x = x
        if y is not None:
            self._y = y
            
        if self.frame and self.frame.winfo_exists():
            self.frame.place(x=self._x, y=max(0, self._y - 24))
            return
        
        try:
            # 创建标签框架
            self.frame = tk.Frame(
                self.parent,
                bg=self.color,
                relief="flat",
                bd=0
            )
            
            # 用户名标签
            self.label = tk.Label(
                self.frame,
                text=f" {self.name} ",
                bg=self.color,
                fg="white",
                font=("Microsoft YaHei UI", 9, "bold"),
                padx=6,
                pady=2
            )
            self.label.pack()
            
            # 定位标签（在光标上方）
            self.frame.place(x=self._x, y=max(0, self._y - 24))
            
            # 启动淡入动画
            self._animate_in()
            
        except Exception:
            pass
    
    def _animate_in(self, step: int = 0):
        """淡入动画"""
        if not self.frame or not self.frame.winfo_exists():
            return
        
        max_steps = 8
        if step < max_steps:
            self._opacity = step / max_steps
            try:
                self.frame.after(30, lambda: self._animate_in(step + 1))
            except Exception:
                pass
    
    def hide(self):
        """隐藏标签"""
        if self._animation_id:
            try:
                self.parent.after_cancel(self._animation_id)
            except Exception:
                pass
        
        if self.frame and self.frame.winfo_exists():
            try:
                self.frame.destroy()
            except Exception:
                pass
        self.frame = None
        self.label = None
    
    def update_position(self, x: int, y: int):
        """更新位置"""
        self._x = x
        self._y = y
        if self.frame and self.frame.winfo_exists():
            try:
                self.frame.place(x=self._x, y=max(0, self._y - 24))
            except Exception:
                pass


class CursorLine:
    """光标竖线 - 闪烁动画"""
    
    def __init__(self, parent, color: str, x: int, y: int, height: int = 20):
        self.parent = parent
        self.color = color
        self.frame: Optional[Any] = None
        self._x = x
        self._y = y
        self._height = height
        self._visible = True
        self._blink_id = None
        
    def show(self, x: int = None, y: int = None):
        """显示光标线"""
        if x is not None:
            self._x = x
        if y is not None:
            self._y = y
            
        if self.frame and self.frame.winfo_exists():
            self.frame.place(x=self._x, y=self._y)
            return
        
        try:
            # 创建细长的竖线
            self.frame = tk.Frame(
                self.parent,
                bg=self.color,
                width=2,
                height=self._height,
                relief="flat",
                bd=0
            )
            self.frame.place(x=self._x, y=self._y)
            
            # 启动闪烁动画
            self._start_blink()
            
        except Exception:
            pass
    
    def _start_blink(self):
        """开始闪烁动画"""
        def blink():
            if not self.frame or not self.frame.winfo_exists():
                return
            
            self._visible = not self._visible
            try:
                if self._visible:
                    self.frame.configure(bg=self.color)
                else:
                    self.frame.configure(bg=self.parent.cget("bg"))
            except Exception:
                pass
            
            self._blink_id = self.parent.after(530, blink)
        
        self._blink_id = self.parent.after(530, blink)
    
    def hide(self):
        """隐藏光标线"""
        if self._blink_id:
            try:
                self.parent.after_cancel(self._blink_id)
            except Exception:
                pass
        
        if self.frame and self.frame.winfo_exists():
            try:
                self.frame.destroy()
            except Exception:
                pass
        self.frame = None
    
    def update_position(self, x: int, y: int, height: int = None):
        """更新位置"""
        self._x = x
        self._y = y
        if height:
            self._height = height
        if self.frame and self.frame.winfo_exists():
            try:
                self.frame.place(x=self._x, y=self._y)
                if height:
                    self.frame.configure(height=height)
            except Exception:
                pass


class CursorManager:
    """远程光标管理器 - 腾讯文档风格增强版
    
    Features:
    - 彩色光标竖线 + 用户名标签
    - 光标闪烁动画
    - 选区高亮（半透明）
    - 不活跃时自动淡出
    """
    
    def __init__(self, editor_widget=None, app=None):
        """初始化光标管理器
        
        Args:
            editor_widget: 编辑器组件
            app: 应用实例（用于调度动画）
        """
        self.editor = editor_widget
        self.app = app
        self.cursors: Dict[str, RemoteCursor] = {}
        self._color_index = 0
        self._cursor_tags: Dict[str, List[str]] = {}
        self._cursor_labels: Dict[str, CursorLabel] = {}
        self._cursor_lines: Dict[str, CursorLine] = {}
        self._activity_check_id = None
        
        # 启动活跃度检查
        self._start_activity_check()

    def add_cursor(self, participant_id: str, name: str, color: str = None) -> str:
        """添加远程光标
        
        Args:
            participant_id: 参与者 ID
            name: 参与者名称
            color: 光标颜色（可选）
            
        Returns:
            分配的颜色
        """
        color_info = self._get_next_color()
        if color is None:
            color = color_info["primary"]
        
        self.cursors[participant_id] = RemoteCursor(
            participant_id=participant_id,
            name=name,
            color=color,
            color_light=color_info.get("light", "#E0E0E0"),
            last_update=time.time()
        )
        
        return color
    
    def remove_cursor(self, participant_id: str) -> None:
        """移除远程光标
        
        Args:
            participant_id: 参与者 ID
        """
        if participant_id in self.cursors:
            # 清除渲染的光标
            self._clear_cursor_display(participant_id)
            del self.cursors[participant_id]
    
    def update_cursor(self, participant_id: str, position: int,
                      selection: Optional[Tuple[int, int]] = None) -> None:
        """更新光标位置
        
        Args:
            participant_id: 参与者 ID
            position: 光标位置
            selection: 选区 (start, end)
        """
        if participant_id not in self.cursors:
            return
        
        cursor = self.cursors[participant_id]
        cursor.position = position
        cursor.last_update = time.time()
        cursor.is_visible = True
        cursor.opacity = 1.0
        
        if selection:
            cursor.selection_start = selection[0]
            cursor.selection_end = selection[1]
        else:
            cursor.selection_start = None
            cursor.selection_end = None
        
        # 重新渲染
        self.render_cursors()
    
    def render_cursors(self) -> None:
        """渲染所有远程光标"""
        if not self.editor:
            return
        
        for participant_id, cursor in self.cursors.items():
            if cursor.is_visible:
                self._draw_cursor(cursor)
                if cursor.has_selection:
                    self._draw_selection(cursor)
    
    def _draw_cursor(self, cursor: RemoteCursor) -> None:
        """绘制单个光标 - 增强版，包含用户名标签和闪烁动画"""
        if not self.editor:
            return
        
        try:
            # 获取文本组件
            text_widget = getattr(self.editor, '_textbox', self.editor)
            
            # 清除旧的光标显示
            self._clear_cursor_display(cursor.participant_id)
            
            # 计算光标位置
            index = f"1.0+{cursor.position}c"
            
            # 获取屏幕坐标
            try:
                bbox = text_widget.bbox(index)
                if bbox:
                    x, y, width, height = bbox
                else:
                    return
            except Exception:
                return
            
            # 创建光标标签
            tag_name = f"cursor_{cursor.participant_id}"
            
            # 创建选区高亮标签
            text_widget.tag_configure(
                tag_name,
                background=cursor.color,
                foreground="white",
                borderwidth=0
            )
            
            # 标记光标位置字符
            try:
                text_widget.tag_add(tag_name, index, f"{index}+1c")
                text_widget.tag_raise(tag_name)
            except Exception:
                pass
            
            # 记录标签
            self._cursor_tags[cursor.participant_id] = [tag_name]
            
            # 创建浮动用户名标签
            if cursor.participant_id not in self._cursor_labels:
                label = CursorLabel(text_widget, cursor.name, cursor.color, x, y)
                self._cursor_labels[cursor.participant_id] = label
            else:
                label = self._cursor_labels[cursor.participant_id]
            
            label.color = cursor.color
            label.name = cursor.name
            label.show(x, y)
            
            # 创建光标竖线
            if cursor.participant_id not in self._cursor_lines:
                line = CursorLine(text_widget, cursor.color, x, y, height)
                self._cursor_lines[cursor.participant_id] = line
            else:
                line = self._cursor_lines[cursor.participant_id]
            
            line.color = cursor.color
            line.show(x, y)
            
        except Exception:
            pass
    
    def _draw_selection(self, cursor: RemoteCursor) -> None:
        """绘制选区高亮 - 半透明渐变效果"""
        if not self.editor or not cursor.has_selection:
            return
        
        try:
            text_widget = getattr(self.editor, '_textbox', self.editor)
            
            # 创建选区标签
            tag_name = f"selection_{cursor.participant_id}"
            
            # 使用浅色背景（更明显但不干扰阅读）
            text_widget.tag_configure(
                tag_name,
                background=cursor.color_light,
                borderwidth=0
            )
            
            # 添加选区
            start_index = f"1.0+{cursor.selection_start}c"
            end_index = f"1.0+{cursor.selection_end}c"
            
            text_widget.tag_add(tag_name, start_index, end_index)
            
            # 确保选区在光标下面
            text_widget.tag_lower(tag_name)
            
            # 记录标签
            if cursor.participant_id in self._cursor_tags:
                self._cursor_tags[cursor.participant_id].append(tag_name)
            else:
                self._cursor_tags[cursor.participant_id] = [tag_name]
            
        except Exception:
            pass
    
    def _clear_cursor_display(self, participant_id: str) -> None:
        """清除光标显示"""
        if not self.editor:
            return
        
        try:
            text_widget = getattr(self.editor, '_textbox', self.editor)
            
            # 移除所有相关标签
            if participant_id in self._cursor_tags:
                for tag_name in self._cursor_tags[participant_id]:
                    try:
                        text_widget.tag_remove(tag_name, "1.0", "end")
                    except Exception:
                        pass
                del self._cursor_tags[participant_id]
            
            # 移除用户名标签
            if participant_id in self._cursor_labels:
                self._cursor_labels[participant_id].hide()
                del self._cursor_labels[participant_id]
            
            # 移除光标竖线
            if participant_id in self._cursor_lines:
                self._cursor_lines[participant_id].hide()
                del self._cursor_lines[participant_id]
            
        except Exception:
            pass
    
    def _get_next_color(self) -> dict:
        """获取下一个可用颜色"""
        color = CURSOR_COLORS[self._color_index % len(CURSOR_COLORS)]
        self._color_index += 1
        return color
    
    def _start_activity_check(self) -> None:
        """启动活跃度检查 - 不活跃的光标自动淡出"""
        def check():
            for participant_id, cursor in list(self.cursors.items()):
                if not cursor.is_active and cursor.is_visible:
                    # 光标不活跃，开始淡出
                    cursor.opacity = max(0, cursor.opacity - 0.2)
                    if cursor.opacity <= 0:
                        cursor.is_visible = False
                        self._fade_out_cursor(participant_id)
            
            if self.app:
                self._activity_check_id = self.app.after(1000, check)
            elif self.editor:
                try:
                    self._activity_check_id = self.editor.after(1000, check)
                except Exception:
                    pass
        
        if self.app:
            self._activity_check_id = self.app.after(1000, check)
        elif self.editor:
            try:
                self._activity_check_id = self.editor.after(1000, check)
            except Exception:
                pass
    
    def _fade_out_cursor(self, participant_id: str) -> None:
        """淡出隐藏光标"""
        if participant_id in self._cursor_labels:
            self._cursor_labels[participant_id].hide()
        if participant_id in self._cursor_lines:
            self._cursor_lines[participant_id].hide()
    
    def clear_all(self) -> None:
        """清除所有光标"""
        for participant_id in list(self.cursors.keys()):
            self.remove_cursor(participant_id)
        
        # 取消活跃度检查
        if self._activity_check_id:
            try:
                if self.app:
                    self.app.after_cancel(self._activity_check_id)
                elif self.editor:
                    self.editor.after_cancel(self._activity_check_id)
            except Exception:
                pass
    
    def get_cursor(self, participant_id: str) -> Optional[RemoteCursor]:
        """获取光标信息"""
        return self.cursors.get(participant_id)
    
    def get_all_cursors(self) -> List[RemoteCursor]:
        """获取所有光标"""
        return list(self.cursors.values())
    
    def get_active_cursors(self) -> List[RemoteCursor]:
        """获取所有活跃光标"""
        return [c for c in self.cursors.values() if c.is_active and c.is_visible]
