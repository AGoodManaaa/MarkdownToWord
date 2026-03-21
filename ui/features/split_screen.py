# -*- coding: utf-8 -*-
"""分屏模式模块 - 支持多种分屏布局"""

import tkinter as tk
from typing import Optional, Callable
from enum import Enum

from ui.theme import save_config  # 用于持久化分屏配置

try:
    import customtkinter as ctk
except ImportError:
    ctk = None


class SplitMode(Enum):
    """分屏模式"""
    HORIZONTAL = "horizontal"  # 左右分屏（默认）
    VERTICAL = "vertical"      # 上下分屏
    EDITOR_ONLY = "editor"     # 仅编辑器
    PREVIEW_ONLY = "preview"   # 仅预览


class SplitScreenManager:
    """分屏管理器"""
    
    def __init__(self, app):
        """
        初始化分屏管理器
        Args:
            app: 主应用实例
        """
        self.app = app
        self._current_mode = SplitMode(app.config.get('split_mode', SplitMode.HORIZONTAL.value))
        self._editor_weight = max(1, int(app.config.get('split_editor_weight', 3)))
        self._preview_weight = max(1, int(app.config.get('split_preview_weight', 2)))
        self._min_pane = 280  # 预览/编辑最小宽度，避免过窄
    
    @property
    def current_mode(self) -> SplitMode:
        """获取当前分屏模式"""
        return self._current_mode
    
    def set_mode(self, mode: SplitMode):
        """设置分屏模式"""
        self._current_mode = mode
        self._apply_mode()
    
    def _apply_mode(self):
        """应用当前分屏模式"""
        mode = self._current_mode
        
        # 获取 PanedWindow 和组件
        paned = getattr(self.app, "paned_window", None)
        input_card = getattr(self.app, "input_card", None)
        preview_card = getattr(self.app, "preview_card", None)
        if paned is None or input_card is None or preview_card is None:
            return
        
        # 清理后重新添加，确保顺序正确
        try:
            for pane in list(paned.panes()):
                paned.forget(pane)
        except Exception:
            pass
        
        if mode == SplitMode.HORIZONTAL:
            paned.configure(orient=tk.HORIZONTAL)
            paned.add(input_card, minsize=self._min_pane, stretch="always")
            paned.add(preview_card, minsize=self._min_pane, stretch="always")
            self.app.preview_visible = True
            self.app.update_status("📐 左右分屏模式")
        
        elif mode == SplitMode.VERTICAL:
            paned.configure(orient=tk.VERTICAL)
            paned.add(input_card, minsize=self._min_pane, stretch="always")
            paned.add(preview_card, minsize=self._min_pane, stretch="always")
            self.app.preview_visible = True
            self.app.update_status("📐 上下分屏模式")
        
        elif mode == SplitMode.EDITOR_ONLY:
            paned.configure(orient=tk.HORIZONTAL)
            paned.add(input_card, minsize=self._min_pane, stretch="always")
            self.app.preview_visible = False
            self.app.update_status("📝 纯编辑模式")
        
        elif mode == SplitMode.PREVIEW_ONLY:
            paned.configure(orient=tk.HORIZONTAL)
            paned.add(preview_card, minsize=self._min_pane, stretch="always")
            self.app.preview_visible = True
            self.app.update_status("👁 纯预览模式")
        
        # 恢复分屏比例
        self._restore_ratio(paned)
        
        # 刷新预览
        if self.app.preview_visible:
            self.app.on_text_change(None)
        
        # 持久化模式
        try:
            self.app.config['split_mode'] = mode.value
            self.app.config['split_editor_weight'] = self._editor_weight
            self.app.config['split_preview_weight'] = self._preview_weight
            save_cfg = getattr(self.app, "save_config", None)
            if callable(save_cfg):
                save_cfg()
        except Exception:
            pass
    
    def toggle_mode(self):
        """循环切换分屏模式"""
        modes = list(SplitMode)
        current_index = modes.index(self._current_mode)
        next_index = (current_index + 1) % len(modes)
        self.set_mode(modes[next_index])
    
    def set_horizontal(self):
        """设置为左右分屏"""
        self.set_mode(SplitMode.HORIZONTAL)
    
    def set_vertical(self):
        """设置为上下分屏"""
        self.set_mode(SplitMode.VERTICAL)
    
    def set_editor_only(self):
        """设置为仅编辑器"""
        self.set_mode(SplitMode.EDITOR_ONLY)
    
    def set_preview_only(self):
        """设置为仅预览"""
        self.set_mode(SplitMode.PREVIEW_ONLY)
    
    def set_ratio(self, editor_weight: int, preview_weight: int):
        """设置分屏比例"""
        self._editor_weight = max(1, editor_weight)
        self._preview_weight = max(1, preview_weight)
        self._apply_mode()
    
    # ---- 内部工具 ----
    def _restore_ratio(self, paned: tk.PanedWindow):
        """按保存的比例恢复分屏"""
        try:
            ratio = float(self.app.config.get('split_ratio', 0.5))
        except Exception:
            ratio = 0.5

        ratio = min(0.85, max(0.15, ratio))

        try:
            if self._current_mode == SplitMode.VERTICAL:
                for delay in (80, 180, 320):
                    paned.after(delay, lambda r=ratio: self._place_vertical_ratio(paned, r))
            elif len(paned.panes()) == 2:
                for delay in (80, 180, 320):
                    paned.after(delay, lambda r=ratio: self._place_horizontal_ratio(paned, r))
        except Exception:
            pass

    def _place_horizontal_ratio(self, paned: tk.PanedWindow, ratio: float):
        try:
            paned.update_idletasks()
            width = max(1, paned.winfo_width())
            self._safe_place_sash(paned, 0, int(width * ratio), 0)
        except Exception:
            pass

    def _place_vertical_ratio(self, paned: tk.PanedWindow, ratio: float):
        try:
            paned.update_idletasks()
            height = max(1, paned.winfo_height())
            self._safe_place_sash(paned, 0, 0, int(height * ratio))
        except Exception:
            pass
    
    def _safe_place_sash(self, paned: tk.PanedWindow, index: int, x: int, y: int):
        """防止异常的 sash_place 封装"""
        try:
            paned.sash_place(index, x, y)
        except Exception:
            pass


class FullScreenPreview:
    """全屏预览功能"""
    
    def __init__(self, app):
        """
        初始化全屏预览
        
        Args:
            app: 主应用实例
        """
        self.app = app
        self._is_fullscreen = False
        self._original_geometry = None
        self._original_state = None
        self._fullscreen_window = None
        self._exit_btn = None
    
    @property
    def is_fullscreen(self) -> bool:
        """是否处于全屏状态"""
        return self._is_fullscreen
    
    def toggle(self):
        """切换全屏状态"""
        if self._is_fullscreen:
            self.exit_fullscreen()
        else:
            self.enter_fullscreen()
    
    def enter_fullscreen(self):
        """进入全屏预览"""
        if self._is_fullscreen:
            return
        
        # 保存原始状态
        self._original_geometry = self.app.geometry()
        self._original_state = self.app.state()
        
        # 创建全屏窗口
        self._fullscreen_window = tk.Toplevel(self.app)
        self._fullscreen_window.title("全屏预览")
        self._fullscreen_window.attributes('-fullscreen', True)
        self._fullscreen_window.configure(bg='white')
        
        # 设置窗口图标
        try:
            import os
            icon_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'app.ico')
            if os.path.exists(icon_path):
                self._fullscreen_window.iconbitmap(icon_path)
        except Exception:
            pass
        
        # 创建预览内容框架
        content_frame = tk.Frame(self._fullscreen_window, bg='white')
        content_frame.pack(fill='both', expand=True, padx=40, pady=40)
        
        # 创建预览文本框
        self._preview_text = tk.Text(
            content_frame,
            wrap='word',
            bg='white',
            fg='#111827',
            font=('宋体', 18),
            padx=40,
            pady=40,
            relief='flat',
            cursor='arrow',
            state='disabled'
        )
        
        # 滚动条
        scrollbar = tk.Scrollbar(content_frame, command=self._preview_text.yview)
        self._preview_text.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side='right', fill='y')
        self._preview_text.pack(side='left', fill='both', expand=True)
        
        # 复制预览内容
        self._copy_preview_content()
        
        # 创建退出按钮（鼠标移到顶部时显示）
        self._exit_frame = tk.Frame(self._fullscreen_window, bg='#1f2937', height=50)
        self._exit_frame.place(relx=0, rely=0, relwidth=1, y=-50)  # 初始隐藏在顶部
        
        exit_btn = tk.Button(
            self._exit_frame,
            text="✕ 退出全屏 (Esc)",
            font=('Microsoft YaHei', 12),
            bg='#ef4444',
            fg='white',
            relief='flat',
            padx=20,
            pady=8,
            cursor='hand2',
            command=self.exit_fullscreen
        )
        exit_btn.pack(pady=10)
        
        # 绑定鼠标移动事件
        self._fullscreen_window.bind('<Motion>', self._on_mouse_move)
        
        # 绑定 Esc 键退出
        self._fullscreen_window.bind('<Escape>', lambda e: self.exit_fullscreen())
        self._fullscreen_window.bind('<F11>', lambda e: self.exit_fullscreen())
        
        self._is_fullscreen = True
        self.app.update_status("🖥️ 全屏预览模式 - 按 Esc 退出")
    
    def _copy_preview_content(self):
        """复制预览内容到全屏窗口"""
        try:
            # 获取原预览区的内容
            content = self.app.preview.text.get('1.0', 'end-1c')
            
            self._preview_text.configure(state='normal')
            self._preview_text.delete('1.0', 'end')
            self._preview_text.insert('1.0', content)
            self._preview_text.configure(state='disabled')
        except Exception as e:
            print(f"复制预览内容失败: {e}")
    
    def _on_mouse_move(self, event):
        """鼠标移动事件 - 显示/隐藏退出按钮"""
        if event.y < 50:
            # 鼠标在顶部，显示退出按钮
            self._exit_frame.place(relx=0, rely=0, relwidth=1, y=0)
        else:
            # 隐藏退出按钮
            self._exit_frame.place(relx=0, rely=0, relwidth=1, y=-50)
    
    def exit_fullscreen(self):
        """退出全屏预览"""
        if not self._is_fullscreen:
            return
        
        # 销毁全屏窗口
        if self._fullscreen_window:
            self._fullscreen_window.destroy()
            self._fullscreen_window = None
        
        self._is_fullscreen = False
        self.app.update_status("👁 已退出全屏预览")


class PrintPreview:
    """打印预览功能"""
    
    def __init__(self, app):
        """
        初始化打印预览
        
        Args:
            app: 主应用实例
        """
        self.app = app
        self._preview_window = None
    
    def show(self):
        """显示打印预览"""
        if self._preview_window and self._preview_window.winfo_exists():
            self._preview_window.focus()
            return
        
        # 创建预览窗口
        self._preview_window = tk.Toplevel(self.app)
        self._preview_window.title("打印预览")
        self._preview_window.geometry("800x900")
        self._preview_window.configure(bg='#f3f4f6')
        
        # 设置窗口图标
        try:
            import os
            icon_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'app.ico')
            if os.path.exists(icon_path):
                self._preview_window.iconbitmap(icon_path)
        except Exception:
            pass
        
        # 工具栏
        toolbar = tk.Frame(self._preview_window, bg='white', height=50)
        toolbar.pack(fill='x', padx=10, pady=10)
        
        # 页面设置
        tk.Label(toolbar, text="页面大小:", bg='white').pack(side='left', padx=(10, 5))
        
        self._page_size_var = tk.StringVar(value="A4")
        page_sizes = ["A4", "A3", "Letter", "Legal"]
        page_combo = tk.OptionMenu(toolbar, self._page_size_var, *page_sizes, command=self._on_page_size_change)
        page_combo.pack(side='left', padx=5)
        
        # 方向
        tk.Label(toolbar, text="方向:", bg='white').pack(side='left', padx=(20, 5))
        
        self._orientation_var = tk.StringVar(value="portrait")
        tk.Radiobutton(toolbar, text="纵向", variable=self._orientation_var, value="portrait", 
                      bg='white', command=self._on_orientation_change).pack(side='left')
        tk.Radiobutton(toolbar, text="横向", variable=self._orientation_var, value="landscape",
                      bg='white', command=self._on_orientation_change).pack(side='left')
        
        # 导出按钮
        export_btn = tk.Button(
            toolbar,
            text="📤 导出 Word",
            font=('Microsoft YaHei', 10),
            bg='#10b981',
            fg='white',
            relief='flat',
            padx=15,
            pady=5,
            cursor='hand2',
            command=self._export
        )
        export_btn.pack(side='right', padx=10)
        
        pdf_btn = tk.Button(
            toolbar,
            text="📄 导出 PDF",
            font=('Microsoft YaHei', 10),
            bg='#3b82f6',
            fg='white',
            relief='flat',
            padx=15,
            pady=5,
            cursor='hand2',
            command=self._export_pdf
        )
        pdf_btn.pack(side='right', padx=5)
        
        # 预览区域（模拟纸张）
        preview_container = tk.Frame(self._preview_window, bg='#f3f4f6')
        preview_container.pack(fill='both', expand=True, padx=20, pady=10)
        
        # 纸张框架
        self._paper_frame = tk.Frame(preview_container, bg='white', relief='solid', bd=1)
        self._paper_frame.pack(expand=True)
        
        # 预览内容
        self._preview_canvas = tk.Canvas(
            self._paper_frame,
            bg='white',
            highlightthickness=0
        )
        self._preview_canvas.pack(fill='both', expand=True, padx=40, pady=40)
        
        # 更新预览
        self._update_preview()
    
    def _on_page_size_change(self, *args):
        """页面大小变化"""
        self._update_preview()
    
    def _on_orientation_change(self):
        """方向变化"""
        self._update_preview()
    
    def _update_preview(self):
        """更新预览"""
        # 获取页面尺寸
        page_size = self._page_size_var.get()
        orientation = self._orientation_var.get()
        
        # 页面尺寸（像素，按 96 DPI）
        sizes = {
            "A4": (595, 842),
            "A3": (842, 1191),
            "Letter": (612, 792),
            "Legal": (612, 1008)
        }
        
        width, height = sizes.get(page_size, (595, 842))
        
        if orientation == "landscape":
            width, height = height, width
        
        # 缩放以适应窗口
        scale = 0.7
        display_width = int(width * scale)
        display_height = int(height * scale)
        
        # 更新纸张大小
        self._paper_frame.configure(width=display_width, height=display_height)
        self._preview_canvas.configure(width=display_width - 80, height=display_height - 80)
        
        # 绘制内容预览
        self._preview_canvas.delete('all')
        
        # 获取预览内容
        try:
            content = self.app.input_text.get('1.0', 'end-1c')
            lines = content.split('\n')[:30]  # 只显示前30行
            
            y = 20
            for line in lines:
                if line.startswith('#'):
                    # 标题
                    level = len(line) - len(line.lstrip('#'))
                    text = line.lstrip('# ')
                    font_size = 16 - level * 2
                    self._preview_canvas.create_text(
                        20, y, text=text, anchor='nw',
                        font=('黑体', font_size, 'bold'),
                        fill='#1f2937'
                    )
                    y += font_size + 10
                else:
                    # 普通文本
                    self._preview_canvas.create_text(
                        20, y, text=line[:50], anchor='nw',
                        font=('宋体', 10),
                        fill='#374151'
                    )
                    y += 16
                
                if y > display_height - 100:
                    self._preview_canvas.create_text(
                        20, y, text="...", anchor='nw',
                        font=('宋体', 10),
                        fill='#9ca3af'
                    )
                    break
        except Exception as e:
            print(f"更新打印预览失败: {e}")
    
    def _export(self):
        """导出 Word"""
        self._preview_window.destroy()
        self.app.export_to_word()
    
    def _export_pdf(self):
        """导出 PDF"""
        self._preview_window.destroy()
        self.app.export_to_pdf()
