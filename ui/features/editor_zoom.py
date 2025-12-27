# -*- coding: utf-8 -*-
"""编辑区缩放功能模块

提供编辑区的缩放控制，支持放大、缩小、重置操作。
支持鼠标滚轮缩放和快捷键缩放。
"""

import customtkinter as ctk
from ui.theme import COLORS, save_config


class EditorZoomFeature:
    """编辑区缩放功能类
    
    提供编辑区的缩放控制：
    - 缩放范围：50% - 150%
    - 步进：10%
    - 支持配置持久化
    - 支持鼠标滚轮缩放（Ctrl+滚轮）
    """
    
    # 缩放范围常量
    MIN_SCALE = 0.5   # 最小 50%
    MAX_SCALE = 1.5   # 最大 150%
    STEP = 0.1        # 步进 10%
    DEFAULT_SCALE = 1.0  # 默认 100%
    BASE_FONT_SIZE = 21  # 基础字体大小（100% 对应 21pt）
    
    def __init__(self, app):
        """初始化编辑区缩放功能
        
        Args:
            app: 主应用实例
        """
        self.app = app
        self._scale = self.DEFAULT_SCALE
        self._scale_label = None
        self._controls_frame = None
        self._bindings_setup = False
    
    def create_controls(self, parent) -> ctk.CTkFrame:
        """创建缩放控件框架
        
        Args:
            parent: 父容器
            
        Returns:
            CTkFrame: 包含缩放控件的框架
        """
        self._controls_frame = ctk.CTkFrame(parent, fg_color="transparent", height=28)
        
        # 缩小按钮
        zoom_out_btn = ctk.CTkButton(
            self._controls_frame,
            text="−",
            width=24,
            height=22,
            corner_radius=6,
            fg_color=COLORS['bg_card'],
            text_color=COLORS['text_primary'],
            hover_color=COLORS['highlight'],
            border_width=1,
            border_color=COLORS['border'],
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.zoom_out
        )
        zoom_out_btn.pack(side="left", padx=1)
        
        # 缩放比例显示
        self._scale_label = ctk.CTkLabel(
            self._controls_frame,
            text="100%",
            font=ctk.CTkFont(size=11),
            text_color=COLORS['text_secondary'],
            width=40
        )
        self._scale_label.pack(side="left", padx=2)
        
        # 放大按钮
        zoom_in_btn = ctk.CTkButton(
            self._controls_frame,
            text="+",
            width=24,
            height=22,
            corner_radius=6,
            fg_color=COLORS['bg_card'],
            text_color=COLORS['text_primary'],
            hover_color=COLORS['highlight'],
            border_width=1,
            border_color=COLORS['border'],
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.zoom_in
        )
        zoom_in_btn.pack(side="left", padx=1)
        
        # 重置按钮
        reset_btn = ctk.CTkButton(
            self._controls_frame,
            text="↺",
            width=24,
            height=22,
            corner_radius=6,
            fg_color=COLORS['bg_card'],
            text_color=COLORS['text_primary'],
            hover_color=COLORS['highlight'],
            border_width=1,
            border_color=COLORS['border'],
            font=ctk.CTkFont(size=12),
            command=self.reset_zoom
        )
        reset_btn.pack(side="left", padx=(4, 0))
        
        # 添加提示
        try:
            self.app.tooltip.add_tooltip(zoom_out_btn, "缩小编辑区")
            self.app.tooltip.add_tooltip(zoom_in_btn, "放大编辑区")
            self.app.tooltip.add_tooltip(reset_btn, "重置缩放")
        except Exception:
            pass
        
        return self._controls_frame
    
    def setup_bindings(self) -> None:
        """设置鼠标滚轮绑定"""
        if self._bindings_setup:
            return
        
        try:
            # 编辑区鼠标滚轮缩放（Ctrl+滚轮）
            if hasattr(self.app, 'input_text') and self.app.input_text:
                self.app.input_text._textbox.bind('<Control-MouseWheel>', self._on_mousewheel)
                self.app.input_text._textbox.bind('<Control-Button-4>', self._on_mousewheel_linux)
                self.app.input_text._textbox.bind('<Control-Button-5>', self._on_mousewheel_linux)
            
            self._bindings_setup = True
        except Exception:
            pass
    
    def _on_mousewheel(self, event) -> str:
        """处理 Windows/Mac 鼠标滚轮事件"""
        try:
            if event.delta > 0:
                self.zoom_in()
            else:
                self.zoom_out()
        except Exception:
            pass
        return 'break'
    
    def _on_mousewheel_linux(self, event) -> str:
        """处理 Linux 鼠标滚轮事件"""
        try:
            if event.num == 4:  # 向上滚动
                self.zoom_in()
            elif event.num == 5:  # 向下滚动
                self.zoom_out()
        except Exception:
            pass
        return 'break'
    
    def zoom_in(self) -> None:
        """放大编辑区（+10%，最大 150%）"""
        new_scale = min(self._scale + self.STEP, self.MAX_SCALE)
        self._set_scale(new_scale)
    
    def zoom_out(self) -> None:
        """缩小编辑区（-10%，最小 50%）"""
        new_scale = max(self._scale - self.STEP, self.MIN_SCALE)
        self._set_scale(new_scale)
    
    def reset_zoom(self) -> None:
        """重置缩放为 100%"""
        self._set_scale(self.DEFAULT_SCALE)
    
    def _set_scale(self, scale: float) -> None:
        """设置缩放比例
        
        Args:
            scale: 缩放比例 (0.5 - 1.5)
        """
        # 限制范围
        scale = max(self.MIN_SCALE, min(self.MAX_SCALE, scale))
        
        # 避免浮点精度问题
        scale = round(scale, 2)
        
        if abs(scale - self._scale) < 0.01:
            return
        
        self._scale = scale
        
        # 更新标签显示
        self._update_label()
        
        # 应用到编辑区（只改字体，不影响布局）
        self._apply_scale()
        
        # 保存配置
        self.save_scale()
        
        # 更新状态栏
        try:
            self.app.update_status(f"🔍 编辑区缩放: {int(scale * 100)}%")
        except Exception:
            pass
    
    def _update_label(self) -> None:
        """更新缩放比例显示标签"""
        if self._scale_label:
            try:
                self._scale_label.configure(text=f"{int(self._scale * 100)}%")
            except Exception:
                pass
    
    def _apply_scale(self) -> None:
        """应用缩放到编辑区（只改字体大小，不影响布局）"""
        try:
            if hasattr(self.app, 'input_text') and self.app.input_text:
                # 计算新字体大小：100% = 21pt
                new_font_size = int(self.BASE_FONT_SIZE * self._scale)
                new_font_size = max(10, min(32, new_font_size))  # 限制范围 10-32pt
                
                # 只更新字体大小，不触发布局重算
                self.app.input_text._textbox.configure(font=('Consolas', new_font_size))
                self.app.input_text.line_numbers.configure(font=('Consolas', new_font_size))
                self.app.input_text.font_size = new_font_size
        except Exception:
            pass
    
    def save_scale(self) -> None:
        """保存缩放比例到配置"""
        try:
            self.app.config['editor_zoom_scale'] = self._scale
            save_config(self.app.config)
        except Exception:
            pass
    
    def restore_scale(self) -> None:
        """从配置恢复缩放比例"""
        try:
            saved_scale = self.app.config.get('editor_zoom_scale', self.DEFAULT_SCALE)
            saved_scale = float(saved_scale)
            saved_scale = max(self.MIN_SCALE, min(self.MAX_SCALE, saved_scale))
            self._scale = saved_scale
            self._update_label()
            self._apply_scale()
        except Exception:
            self._scale = self.DEFAULT_SCALE
        
        # 设置绑定
        self.setup_bindings()
    
    @property
    def scale(self) -> float:
        """获取当前缩放比例"""
        return self._scale
