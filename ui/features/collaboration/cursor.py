# -*- coding: utf-8 -*-
"""远程光标管理模块"""

from dataclasses import dataclass
from typing import Dict, Optional, Tuple, List
import time


@dataclass
class RemoteCursor:
    """远程光标"""
    participant_id: str
    name: str
    color: str
    position: int = 0
    selection_start: Optional[int] = None
    selection_end: Optional[int] = None
    last_update: float = 0.0
    
    @property
    def has_selection(self) -> bool:
        """是否有选区"""
        return self.selection_start is not None and self.selection_end is not None


class CursorManager:
    """远程光标管理器"""
    
    CURSOR_COLORS = [
        "#FF6B6B", "#4ECDC4", "#45B7D1", "#96CEB4",
        "#FFEAA7", "#DDA0DD", "#98D8C8", "#F7DC6F"
    ]
    
    def __init__(self, editor_widget=None):
        """初始化光标管理器
        
        Args:
            editor_widget: 编辑器组件
        """
        self.editor = editor_widget
        self.cursors: Dict[str, RemoteCursor] = {}
        self._color_index = 0
        self._cursor_tags: Dict[str, List[str]] = {}

    def add_cursor(self, participant_id: str, name: str, color: str = None) -> str:
        """添加远程光标
        
        Args:
            participant_id: 参与者 ID
            name: 参与者名称
            color: 光标颜色（可选）
            
        Returns:
            分配的颜色
        """
        if color is None:
            color = self._get_next_color()
        
        self.cursors[participant_id] = RemoteCursor(
            participant_id=participant_id,
            name=name,
            color=color,
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
            self._draw_cursor(cursor)
            if cursor.has_selection:
                self._draw_selection(cursor)
    
    def _draw_cursor(self, cursor: RemoteCursor) -> None:
        """绘制单个光标"""
        if not self.editor:
            return
        
        try:
            # 获取文本组件
            text_widget = getattr(self.editor, '_textbox', self.editor)
            
            # 清除旧的光标显示
            self._clear_cursor_display(cursor.participant_id)
            
            # 计算光标位置
            index = f"1.0+{cursor.position}c"
            
            # 创建光标标签
            tag_name = f"cursor_{cursor.participant_id}"
            
            # 配置标签样式
            text_widget.tag_configure(
                tag_name,
                background=cursor.color,
                foreground="white"
            )
            
            # 在光标位置添加标记（使用一个特殊字符或边框）
            # 由于 tkinter 限制，我们使用背景色标记
            try:
                text_widget.tag_add(tag_name, index, f"{index}+1c")
            except Exception:
                pass
            
            # 记录标签
            self._cursor_tags[cursor.participant_id] = [tag_name]
            
        except Exception:
            pass
    
    def _draw_selection(self, cursor: RemoteCursor) -> None:
        """绘制选区高亮"""
        if not self.editor or not cursor.has_selection:
            return
        
        try:
            text_widget = getattr(self.editor, '_textbox', self.editor)
            
            # 创建选区标签
            tag_name = f"selection_{cursor.participant_id}"
            
            # 配置半透明背景
            text_widget.tag_configure(
                tag_name,
                background=cursor.color + "40"  # 添加透明度
            )
            
            # 添加选区
            start_index = f"1.0+{cursor.selection_start}c"
            end_index = f"1.0+{cursor.selection_end}c"
            
            text_widget.tag_add(tag_name, start_index, end_index)
            
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
            
        except Exception:
            pass
    
    def _get_next_color(self) -> str:
        """获取下一个可用颜色"""
        color = self.CURSOR_COLORS[self._color_index % len(self.CURSOR_COLORS)]
        self._color_index += 1
        return color
    
    def clear_all(self) -> None:
        """清除所有光标"""
        for participant_id in list(self.cursors.keys()):
            self.remove_cursor(participant_id)
    
    def get_cursor(self, participant_id: str) -> Optional[RemoteCursor]:
        """获取光标信息"""
        return self.cursors.get(participant_id)
    
    def get_all_cursors(self) -> List[RemoteCursor]:
        """获取所有光标"""
        return list(self.cursors.values())
