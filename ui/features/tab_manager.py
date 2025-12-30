# -*- coding: utf-8 -*-
"""多标签页编辑功能模块

支持在同一窗口中编辑多个 Markdown 文件。
"""

import uuid
import tkinter as tk
from tkinter import messagebox
from dataclasses import dataclass, field
from typing import List, Optional, Callable
import customtkinter as ctk

from ui.theme import COLORS


@dataclass
class TabData:
    """标签页数据类"""
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    title: str = "未命名"
    file_path: Optional[str] = None
    content: str = ""
    modified: bool = False
    cursor_position: str = "1.0"
    scroll_position: float = 0.0


class TabManagerFeature:
    """多标签页管理功能类
    
    提供标签页式多文档编辑：
    - 创建/关闭/切换标签页
    - 未保存状态管理
    - 右键菜单操作
    """
    
    MAX_TABS = 20  # 最大标签页数量
    
    def __init__(self, app):
        """初始化标签页管理器
        
        Args:
            app: 主应用实例
        """
        self.app = app
        self.tabs: List[TabData] = []
        self.active_tab_id: Optional[str] = None
        self.tab_bar: Optional[ctk.CTkFrame] = None
        self._tab_buttons: dict = {}  # tab_id -> button frame
        self._tabs_container: Optional[ctk.CTkFrame] = None
    
    def create_tab_bar(self, parent) -> ctk.CTkFrame:
        """创建标签栏
        
        Args:
            parent: 父容器
            
        Returns:
            CTkFrame: 标签栏框架
        """
        self.tab_bar = ctk.CTkFrame(parent, fg_color=COLORS['bg_sidebar'], height=32, corner_radius=0)
        
        # 标签页容器（可滚动）
        self._tabs_container = ctk.CTkFrame(self.tab_bar, fg_color="transparent")
        self._tabs_container.pack(side="left", fill="x", expand=True, padx=(4, 0))
        
        # 新建标签按钮
        new_tab_btn = ctk.CTkButton(
            self.tab_bar,
            text="+",
            width=28,
            height=24,
            corner_radius=6,
            fg_color="transparent",
            text_color=COLORS['text_secondary'],
            hover_color=COLORS['highlight'],
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.new_tab
        )
        new_tab_btn.pack(side="right", padx=4, pady=4)
        
        try:
            self.app.tooltip.add_tooltip(new_tab_btn, "新建标签页")
        except Exception:
            pass
        
        # 创建初始标签页
        if not self.tabs:
            self.new_tab()
        
        return self.tab_bar
    
    def new_tab(self, file_path: Optional[str] = None, content: str = "") -> str:
        """创建新标签页
        
        Args:
            file_path: 文件路径（可选）
            content: 初始内容
            
        Returns:
            str: 新标签页的 ID
        """
        if len(self.tabs) >= self.MAX_TABS:
            messagebox.showwarning("提示", f"最多只能打开 {self.MAX_TABS} 个标签页")
            return self.active_tab_id or ""
        
        # 确定标题
        if file_path:
            import os
            title = os.path.basename(file_path)
        else:
            # 计算未命名标签的编号
            untitled_count = sum(1 for t in self.tabs if t.title.startswith("未命名"))
            title = f"未命名{untitled_count + 1}" if untitled_count > 0 else "未命名"
        
        # 创建标签数据
        tab = TabData(
            title=title,
            file_path=file_path,
            content=content,
            modified=False
        )
        self.tabs.append(tab)
        
        # 创建标签按钮
        self._create_tab_button(tab)
        
        # 切换到新标签
        self.switch_tab(tab.id)
        
        return tab.id
    
    def _create_tab_button(self, tab_data: TabData) -> ctk.CTkFrame:
        """创建单个标签按钮
        
        Args:
            tab_data: 标签数据
            
        Returns:
            CTkFrame: 标签按钮框架
        """
        # 标签框架
        tab_frame = ctk.CTkFrame(
            self._tabs_container,
            fg_color=COLORS['bg_card'],
            corner_radius=6,
            height=26
        )
        tab_frame.pack(side="left", padx=2, pady=3)
        
        # 标签标题
        title_text = tab_data.title
        if tab_data.modified:
            title_text = f"● {title_text}"
        
        title_label = ctk.CTkLabel(
            tab_frame,
            text=title_text,
            font=ctk.CTkFont(size=11),
            text_color=COLORS['text_primary'],
            cursor="hand2"
        )
        title_label.pack(side="left", padx=(8, 4), pady=2)
        
        # 关闭按钮
        close_btn = ctk.CTkButton(
            tab_frame,
            text="×",
            width=18,
            height=18,
            corner_radius=4,
            fg_color="transparent",
            text_color=COLORS['text_secondary'],
            hover_color=COLORS['danger'],
            font=ctk.CTkFont(size=12),
            command=lambda: self.close_tab(tab_data.id)
        )
        close_btn.pack(side="right", padx=(0, 4), pady=2)
        
        # 绑定点击事件
        title_label.bind('<Button-1>', lambda e: self.switch_tab(tab_data.id))
        tab_frame.bind('<Button-1>', lambda e: self.switch_tab(tab_data.id))
        
        # 绑定右键菜单
        title_label.bind('<Button-3>', lambda e: self._show_context_menu(e, tab_data.id))
        tab_frame.bind('<Button-3>', lambda e: self._show_context_menu(e, tab_data.id))
        
        # 存储引用
        self._tab_buttons[tab_data.id] = {
            'frame': tab_frame,
            'label': title_label,
            'close_btn': close_btn
        }
        
        return tab_frame
    
    def close_tab(self, tab_id: str) -> bool:
        """关闭标签页
        
        Args:
            tab_id: 标签页 ID
            
        Returns:
            bool: 是否成功关闭
        """
        tab = self._get_tab_by_id(tab_id)
        if not tab:
            return False
        
        # 检查未保存更改
        if tab.modified:
            result = messagebox.askyesnocancel(
                "未保存的更改",
                f'"{tab.title}" 有未保存的更改。\n是否保存？'
            )
            if result is None:  # 取消
                return False
            if result:  # 保存
                # 保存当前标签内容
                if tab_id == self.active_tab_id:
                    self._save_current_tab_state()
                if not self._save_tab(tab):
                    return False
        
        # 移除标签按钮
        if tab_id in self._tab_buttons:
            self._tab_buttons[tab_id]['frame'].destroy()
            del self._tab_buttons[tab_id]
        
        # 移除标签数据
        self.tabs = [t for t in self.tabs if t.id != tab_id]
        
        # 如果关闭的是当前标签，切换到其他标签
        if tab_id == self.active_tab_id:
            if self.tabs:
                self.switch_tab(self.tabs[-1].id)
            else:
                # 所有标签都关闭了，创建新空白标签
                self.active_tab_id = None
                self.new_tab()
        
        return True
    
    def switch_tab(self, tab_id: str) -> None:
        """切换到指定标签页
        
        Args:
            tab_id: 目标标签页 ID
        """
        if tab_id == self.active_tab_id:
            return
        
        tab = self._get_tab_by_id(tab_id)
        if not tab:
            return
        
        # 保存当前标签状态
        if self.active_tab_id:
            self._save_current_tab_state()
        
        # 更新活动标签
        self.active_tab_id = tab_id
        
        # 更新标签按钮样式
        self._update_tab_styles()
        
        # 恢复目标标签内容
        self._restore_tab_content(tab)
        
        # 更新应用状态
        self.app.current_file = tab.file_path
        self.app._content_modified = tab.modified
        self.app._update_title()
    
    def _save_current_tab_state(self) -> None:
        """保存当前标签页状态"""
        if not self.active_tab_id:
            return
        
        tab = self._get_tab_by_id(self.active_tab_id)
        if not tab:
            return
        
        try:
            # 保存内容
            tab.content = self.app.input_text.get("1.0", "end-1c")
            
            # 保存光标位置
            tab.cursor_position = self.app.input_text._textbox.index("insert")
            
            # 保存滚动位置
            tab.scroll_position = self.app.input_text._textbox.yview()[0]
            
            # 保存修改状态
            tab.modified = self.app._content_modified
        except Exception:
            pass
    
    def _restore_tab_content(self, tab: TabData) -> None:
        """恢复标签页内容
        
        Args:
            tab: 标签数据
        """
        try:
            # 清空并设置内容
            self.app.input_text.delete("1.0", "end")
            if tab.content:
                self.app.input_text.insert("1.0", tab.content)
            
            # 恢复光标位置
            self.app.input_text._textbox.mark_set("insert", tab.cursor_position)
            
            # 恢复滚动位置
            self.app.input_text._textbox.yview_moveto(tab.scroll_position)
            
            # 更新预览
            self.app.on_text_change(None)
            
            # 更新修改状态
            self.app._content_modified = tab.modified
            self.app._last_saved_content = tab.content if not tab.modified else ""
        except Exception:
            pass
    
    def _update_tab_styles(self) -> None:
        """更新所有标签按钮样式"""
        for tab_id, btn_info in self._tab_buttons.items():
            is_active = tab_id == self.active_tab_id
            try:
                btn_info['frame'].configure(
                    fg_color=COLORS['primary'] if is_active else COLORS['bg_card']
                )
                btn_info['label'].configure(
                    text_color="white" if is_active else COLORS['text_primary']
                )
                btn_info['close_btn'].configure(
                    text_color="white" if is_active else COLORS['text_secondary'],
                    hover_color=COLORS['primary_hover'] if is_active else COLORS['danger']
                )
            except Exception:
                pass
    
    def update_tab_title(self, tab_id: str, title: str = None, modified: bool = None) -> None:
        """更新标签页标题
        
        Args:
            tab_id: 标签页 ID
            title: 新标题（可选）
            modified: 修改状态（可选）
        """
        tab = self._get_tab_by_id(tab_id)
        if not tab:
            return
        
        if title is not None:
            tab.title = title
        if modified is not None:
            tab.modified = modified
        
        # 更新显示
        if tab_id in self._tab_buttons:
            display_title = tab.title
            if tab.modified:
                display_title = f"● {display_title}"
            try:
                self._tab_buttons[tab_id]['label'].configure(text=display_title)
            except Exception:
                pass
    
    def get_active_tab(self) -> Optional[TabData]:
        """获取当前活动标签页
        
        Returns:
            Optional[TabData]: 当前标签数据，无则返回 None
        """
        return self._get_tab_by_id(self.active_tab_id) if self.active_tab_id else None
    
    def _get_tab_by_id(self, tab_id: str) -> Optional[TabData]:
        """根据 ID 获取标签数据"""
        for tab in self.tabs:
            if tab.id == tab_id:
                return tab
        return None
    
    def _save_tab(self, tab: TabData) -> bool:
        """保存标签页内容到文件"""
        if not tab.file_path:
            # 需要另存为
            from tkinter import filedialog
            file_path = filedialog.asksaveasfilename(
                title="保存文件",
                defaultextension=".md",
                filetypes=[("Markdown 文件", "*.md"), ("所有文件", "*.*")]
            )
            if not file_path:
                return False
            tab.file_path = file_path
            import os
            tab.title = os.path.basename(file_path)
        
        try:
            with open(tab.file_path, 'w', encoding='utf-8') as f:
                f.write(tab.content)
            tab.modified = False
            self.update_tab_title(tab.id)
            return True
        except Exception as e:
            messagebox.showerror("保存失败", f"无法保存文件:\n{str(e)}")
            return False
    
    def _show_context_menu(self, event, tab_id: str) -> None:
        """显示标签页右键菜单
        
        Args:
            event: 鼠标事件
            tab_id: 标签页 ID
        """
        menu = tk.Menu(self.app, tearoff=0)
        menu.add_command(label="关闭", command=lambda: self.close_tab(tab_id))
        menu.add_command(label="关闭其他", command=lambda: self._close_other_tabs(tab_id))
        menu.add_command(label="关闭全部", command=self._close_all_tabs)
        
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
    
    def _close_other_tabs(self, keep_tab_id: str) -> None:
        """关闭除指定标签外的所有标签"""
        tabs_to_close = [t.id for t in self.tabs if t.id != keep_tab_id]
        for tab_id in tabs_to_close:
            if not self.close_tab(tab_id):
                break  # 用户取消了
    
    def _close_all_tabs(self) -> None:
        """关闭所有标签"""
        tabs_to_close = [t.id for t in self.tabs]
        for tab_id in tabs_to_close:
            if not self.close_tab(tab_id):
                break
    
    def mark_current_modified(self) -> None:
        """标记当前标签为已修改"""
        if self.active_tab_id:
            self.update_tab_title(self.active_tab_id, modified=True)
    
    def open_file_in_tab(self, file_path: str, content: str) -> str:
        """在新标签页中打开文件
        
        Args:
            file_path: 文件路径
            content: 文件内容
            
        Returns:
            str: 标签页 ID
        """
        # 检查是否已打开
        for tab in self.tabs:
            if tab.file_path == file_path:
                self.switch_tab(tab.id)
                return tab.id
        
        return self.new_tab(file_path=file_path, content=content)

    def save_tab_as(self, tab: TabData) -> bool:
        """另存为标签页内容
        
        Args:
            tab: 标签数据
            
        Returns:
            bool: 是否保存成功
        """
        from tkinter import filedialog
        file_path = filedialog.asksaveasfilename(
            title="另存为",
            defaultextension=".md",
            filetypes=[("Markdown 文件", "*.md"), ("所有文件", "*.*")]
        )
        if not file_path:
            return False
        
        tab.file_path = file_path
        import os
        tab.title = os.path.basename(file_path)
        
        return self._save_tab(tab)
    
    def check_all_tabs_unsaved_changes(self) -> bool:
        """检查所有标签页是否有未保存的更改
        
        Returns:
            bool: 如果所有更改都处理了（保存或不保存）则返回 True，取消则返回 False
        """
        for tab in self.tabs:
            if tab.modified:
                # 切换到该标签让用户看到
                self.switch_tab(tab.id)
                result = messagebox.askyesnocancel(
                    "未保存的更改",
                    f'"{tab.title}" 有未保存的更改。\n是否保存？'
                )
                if result is None:  # 取消
                    return False
                if result:  # 保存
                    # 确保内容是最新的
                    if tab.id == self.active_tab_id:
                        self._save_current_tab_state()
                    if not self._save_tab(tab):
                        return False
                else:
                    # 不保存，恢复修改状态以允许关闭
                    pass
        return True
