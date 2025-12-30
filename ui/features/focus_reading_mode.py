# -*- coding: utf-8 -*-
"""
专注模式和阅读模式功能模块
"""

import customtkinter as ctk
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from gui import App


from ui.theme import COLORS

class FocusModeFeature:
    """专注模式 - 隐藏所有干扰元素，专注于编辑"""
    
    def __init__(self, app: 'App'):
        self.app = app
        self.is_active = False
        self._saved_state = {}
        self._saved_widgets = {}  # 保存隐藏的widgets引用
        self.exit_btn = None
    
    def toggle(self):
        """切换专注模式"""
        if self.is_active:
            self.exit()
        else:
            self.enter()
    
    def enter(self):
        """进入专注模式"""
        if self.is_active:
            return
        
        # 保存当前状态
        self._saved_state = {
            'sidebar_visible': hasattr(self.app, 'sidebar_visible') and self.app.sidebar_visible,
            'preview_visible': hasattr(self.app, 'preview_visible') and self.app.preview_visible,
        }
        
        self._saved_widgets = {}
        
        # 1. 隐藏侧边栏
        if hasattr(self.app, 'sidebar') and self.app.sidebar.winfo_ismapped():
            self._saved_widgets['sidebar'] = True
            self.app.sidebar.pack_forget()
            if hasattr(self.app, 'sidebar_visible'):
                self.app.sidebar_visible = False
        
        # 2. 隐藏预览区
        if hasattr(self.app, 'preview_card') and self.app.preview_card.winfo_ismapped():
            self._saved_widgets['preview'] = True
            self.app.preview_card.grid_forget()
            if hasattr(self.app, 'preview_visible'):
                self.app.preview_visible = False
        
        # 3. 布局调整：专注模式 - 增加宽度到 1200
        self.app.main_frame.grid_columnconfigure(0, weight=1)
        self.app.main_frame.grid_columnconfigure(1, weight=0, minsize=1200)
        self.app.main_frame.grid_columnconfigure(2, weight=1)
        
        # 将编辑区移到中间列
        if hasattr(self.app, 'input_card'):
            self._saved_widgets['input_pos'] = True # 标记位置改变
            self.app.input_card.grid(row=0, column=1, sticky="nsew", padx=40, pady=20)
            
        # 4. 隐藏中间工具栏 (智能判断：不隐藏右侧pack的frame)
        if hasattr(self.app, 'header'):
            for child in self.app.header.winfo_children():
                if isinstance(child, ctk.CTkFrame) and child.winfo_ismapped():
                    # 检查pack side，跳过右侧的（保留功能按钮）
                    try:
                        info = child.pack_info()
                        if info.get('side') == 'right':
                            continue
                    except:
                        pass
                    
                    # 隐藏左侧/中间的大工具栏
                    # 通常左侧是标题(children=1)，中间是工具(children>5)
                    if len(child.winfo_children()) > 3:
                        self._saved_widgets[f'toolbar_{id(child)}'] = child
                        child.pack_forget()
        
        # 5. 隐藏状态栏
        if hasattr(self.app, 'status_bar') and self.app.status_bar.winfo_ismapped():
            self._saved_widgets['status_bar'] = True
            self.app.status_bar.pack_forget()
        
        self.is_active = True
        
        # 6. 添加浮动退出按钮
        try:
            self.exit_btn = ctk.CTkButton(
                self.app,
                text="退出专注模式",
                command=self.exit,
                fg_color=COLORS['primary'],
                font=("Microsoft YaHei", 12, "bold"),
                width=140,
                height=45,
                corner_radius=22,
                border_width=2,
                border_color="white",
                alpha=0.9
            )
            # 放置在右下角
            self.exit_btn.place(relx=0.95, rely=0.95, anchor="se")
            self.exit_btn.lift() # 确保在最顶层
        except:
            pass
        
        # 安全更新状态
        try:
            self.app.update_status("🎯 专注模式已启用 - 按 F11 退出")
        except:
            pass
        
        # 更新按钮状态
        if hasattr(self.app, 'header_styler'):
            try:
                self.app.header_styler.update_states()
            except:
                pass

        # 更新专注模式按钮高亮
        try:
            if hasattr(self.app, 'focus_mode_btn'):
                self.app.focus_mode_btn.configure(fg_color=COLORS['primary_hover'])
        except:
            pass
    
    def exit(self):
        """退出专注模式"""
        if not self.is_active:
            return
            
        # 移除浮动按钮
        if self.exit_btn:
            try:
                self.exit_btn.destroy()
                self.exit_btn = None
            except:
                pass
        
        # 恢复侧边栏
        if self._saved_widgets.get('sidebar'):
            if hasattr(self.app, 'sidebar'):
                self.app.sidebar.pack(side="left", fill="y", padx=(0, 10), before=self.app.right_container if hasattr(self.app, 'right_container') else None)
                if hasattr(self.app, 'sidebar_visible'):
                    self.app.sidebar_visible = True
        
        # 恢复编辑区到左侧
        if hasattr(self.app, 'input_card'):
            self.app.input_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
            
        # 恢复预览区
        if self._saved_widgets.get('preview'):
            if hasattr(self.app, 'preview_card'):
                self.app.preview_card.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
                if hasattr(self.app, 'preview_visible'):
                    self.app.preview_visible = True
        
        # 恢复列权重
        self.app.main_frame.grid_columnconfigure(0, weight=3)
        self.app.main_frame.grid_columnconfigure(1, weight=2)
        self.app.main_frame.grid_columnconfigure(2, weight=0)
        
        # 恢复工具栏
        if hasattr(self.app, 'header'):
            for key, widget in self._saved_widgets.items():
                if key.startswith('toolbar_') and isinstance(widget, ctk.CTkFrame):
                    if not widget.winfo_ismapped():
                        # 尝试恢复到标题后面
                        try:
                            if hasattr(self.app, 'title_label') and self.app.title_label.master:
                                widget.pack(side="left", padx=24, after=self.app.title_label.master)
                            else:
                                widget.pack(side="left", padx=24)
                        except:
                            widget.pack(side="left", padx=24)
        
        # 恢复状态栏
        if self._saved_widgets.get('status_bar'):
            if hasattr(self.app, 'status_bar') and not self.app.status_bar.winfo_ismapped():
                self.app.status_bar.pack(side="bottom", fill="x")
        
        self.is_active = False
        
        # 更新状态
        try:
            self.app.update_status("✨ 已退出专注模式")
        except:
            pass
        
        # 更新按钮状态
        if hasattr(self.app, 'header_styler'):
            try:
                self.app.header_styler.update_states()
            except:
                pass
                
        # 恢复按钮样式
        try:
            if hasattr(self.app, 'focus_mode_btn'):
                self.app.focus_mode_btn.configure(fg_color="transparent")
        except:
            pass


class ReadingModeFeature:
    """阅读模式 - 隐藏编辑器，优化预览区显示（Typora风格）"""
    
    def __init__(self, app: 'App'):
        self.app = app
        self.is_active = False
        self._saved_state = {}
        self._saved_widgets = {}
        self._reading_container = None
        self.reading_width = 1800 # 增加到 1800
        self.exit_btn = None
    
    def toggle(self):
        """切换阅读模式"""
        if self.is_active:
            self.exit()
        else:
            self.enter()
    
    def enter(self):
        """进入阅读模式"""
        if self.is_active:
            return
        
        self._saved_state = {
            'sidebar_visible': hasattr(self.app, 'sidebar_visible') and self.app.sidebar_visible,
        }
        self._saved_widgets = {}
        
        # 1. 隐藏侧边栏
        if hasattr(self.app, 'sidebar') and self.app.sidebar.winfo_ismapped():
            self._saved_widgets['sidebar'] = True
            self.app.sidebar.pack_forget()
            if hasattr(self.app, 'sidebar_visible'):
                self.app.sidebar_visible = False
        
        # 2. 隐藏编辑区
        if hasattr(self.app, 'input_card') and self.app.input_card.winfo_ismapped():
            self._saved_widgets['input'] = True
            self.app.input_card.grid_forget()
        
        # 3. 隐藏中间工具栏
        if hasattr(self.app, 'header'):
            for child in self.app.header.winfo_children():
                if isinstance(child, ctk.CTkFrame) and child.winfo_ismapped():
                    # 检查pack side，跳过右侧的
                    try:
                        info = child.pack_info()
                        if info.get('side') == 'right':
                            continue
                    except:
                        pass
                    
                    # 隐藏左侧/中间的大工具栏
                    if len(child.winfo_children()) > 5:
                        self._saved_widgets[f'toolbar_{id(child)}'] = child
                        child.pack_forget()

        # 4. 隐藏状态栏
        if hasattr(self.app, 'status_bar') and self.app.status_bar.winfo_ismapped():
            self._saved_widgets['status_bar'] = True
            self.app.status_bar.pack_forget()
            
        # 5. 布局调整 - 宽度 1400
        self.app.main_frame.grid_columnconfigure(0, weight=1)
        self.app.main_frame.grid_columnconfigure(1, weight=0, minsize=self.reading_width)
        self.app.main_frame.grid_columnconfigure(2, weight=1)
        
        if hasattr(self.app, 'preview_card'):
            self.app.preview_card.grid(row=0, column=1, sticky="nsew", padx=40, pady=20)
            if hasattr(self.app, 'preview_visible'):
                self.app.preview_visible = True
                
        self.is_active = True
        
        # 6. 添加浮动退出按钮
        try:
            self.exit_btn = ctk.CTkButton(
                self.app,
                text="退出阅读模式",
                command=self.exit,
                fg_color=COLORS['primary'],
                font=("Microsoft YaHei", 12, "bold"),
                width=140,
                height=45,
                corner_radius=22,
                border_width=2,
                border_color="white",
                alpha=0.9
            )
            # 放置在右下角
            self.exit_btn.place(relx=0.95, rely=0.95, anchor="se")
            self.exit_btn.lift()
        except:
            pass
            
        # 更新按钮高亮
        try:
            if hasattr(self.app, 'reading_mode_btn'):
                self.app.reading_mode_btn.configure(fg_color="#16A34A")
        except:
            pass

    def exit(self):
        """退出阅读模式"""
        if not self.is_active:
            return
            
        # 移除浮动按钮
        if self.exit_btn:
            try:
                self.exit_btn.destroy()
                self.exit_btn = None
            except:
                pass

        # 恢复布局
        self.app.main_frame.grid_columnconfigure(0, weight=3)
        self.app.main_frame.grid_columnconfigure(1, weight=2)
        self.app.main_frame.grid_columnconfigure(2, weight=0)
        
        # 恢复侧边栏
        if self._saved_widgets.get('sidebar'):
            if hasattr(self.app, 'sidebar'):
                self.app.sidebar.pack(side="left", fill="y", padx=(0, 10), before=self.app.right_container if hasattr(self.app, 'right_container') else None)
                if hasattr(self.app, 'sidebar_visible'):
                    self.app.sidebar_visible = True
        
        # 恢复编辑区
        if self._saved_widgets.get('input'):
            if hasattr(self.app, 'input_card'):
                self.app.input_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        
        # 恢复预览区
        if hasattr(self.app, 'preview_card'):
            self.app.preview_card.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
            
        # 恢复工具栏
        if hasattr(self.app, 'header'):
            for key, widget in self._saved_widgets.items():
                if key.startswith('toolbar_') and isinstance(widget, ctk.CTkFrame):
                    if not widget.winfo_ismapped():
                        # 尝试恢复到标题后面
                        try:
                            if hasattr(self.app, 'title_label') and self.app.title_label.master:
                                widget.pack(side="left", padx=24, after=self.app.title_label.master)
                            else:
                                widget.pack(side="left", padx=24)
                        except:
                            widget.pack(side="left", padx=24)
                    
        # 恢复状态栏
        if self._saved_widgets.get('status_bar'):
            if hasattr(self.app, 'status_bar'):
                self.app.status_bar.pack(side="bottom", fill="x")
        
        self.is_active = False
        
        # 恢复按钮样式
        try:
            if hasattr(self.app, 'reading_mode_btn'):
                self.app.reading_mode_btn.configure(fg_color="transparent")
        except:
            pass
        
        try:
            self.app.update_status("✨ 已退出阅读模式")
        except:
            pass
