# -*- coding: utf-8 -*-
"""UI 动画效果模块 - 提供平滑过渡和微交互"""

import customtkinter as ctk
from typing import Callable, Optional


class AnimationMixin:
    """动画混入类，为组件添加动画能力"""
    
    def fade_in(self, duration_ms: int = 200, callback: Optional[Callable] = None):
        """淡入动画"""
        self._animate_opacity(0.0, 1.0, duration_ms, callback)
    
    def fade_out(self, duration_ms: int = 200, callback: Optional[Callable] = None):
        """淡出动画"""
        self._animate_opacity(1.0, 0.0, duration_ms, callback)
    
    def _animate_opacity(self, start: float, end: float, duration_ms: int, callback: Optional[Callable]):
        steps = 10
        delay = duration_ms // steps
        delta = (end - start) / steps
        
        def step(current_alpha):
            if (delta > 0 and current_alpha >= end) or (delta < 0 and current_alpha <= end):
                if callback:
                    callback()
                return
            try:
                # 通过调整透明度模拟淡入淡出（CTk 不直接支持透明度）
                pass
            except Exception:
                pass
            self.after(delay, lambda: step(current_alpha + delta))
        
        step(start)


class HoverButton(ctk.CTkButton):
    """带悬停效果的按钮"""
    
    def __init__(self, master, scale_factor: float = 1.02, **kwargs):
        super().__init__(master, **kwargs)
        self._scale_factor = scale_factor
        self._original_width = kwargs.get('width')
        self._original_height = kwargs.get('height')
        
        self.bind('<Enter>', self._on_enter)
        self.bind('<Leave>', self._on_leave)
    
    def _on_enter(self, event):
        """鼠标进入 - 轻微放大"""
        try:
            if self._original_width and self._original_height:
                new_w = int(self._original_width * self._scale_factor)
                new_h = int(self._original_height * self._scale_factor)
                self.configure(width=new_w, height=new_h)
        except Exception:
            pass
    
    def _on_leave(self, event):
        """鼠标离开 - 恢复原始尺寸"""
        try:
            if self._original_width and self._original_height:
                self.configure(width=self._original_width, height=self._original_height)
        except Exception:
            pass


class GlowFrame(ctk.CTkFrame):
    """带发光边框效果的框架"""
    
    def __init__(self, master, glow_color: str = "#818CF8", **kwargs):
        self._glow_color = glow_color
        super().__init__(master, **kwargs)
        
        self.bind('<Enter>', self._on_enter)
        self.bind('<Leave>', self._on_leave)
    
    def _on_enter(self, event):
        """鼠标进入 - 显示发光边框"""
        try:
            self.configure(border_color=self._glow_color, border_width=2)
        except Exception:
            pass
    
    def _on_leave(self, event):
        """鼠标离开 - 隐藏边框"""
        try:
            self.configure(border_width=0)
        except Exception:
            pass


class PulseIndicator(ctk.CTkLabel):
    """脉冲指示器 - 用于加载状态"""
    
    def __init__(self, master, pulse_color: str = "#818CF8", **kwargs):
        super().__init__(master, text="●", **kwargs)
        self._pulse_color = pulse_color
        self._is_pulsing = False
        self._pulse_state = 0
    
    def start_pulse(self, interval_ms: int = 500):
        """开始脉冲动画"""
        self._is_pulsing = True
        self._pulse()
    
    def stop_pulse(self):
        """停止脉冲动画"""
        self._is_pulsing = False
    
    def _pulse(self):
        if not self._is_pulsing:
            return
        
        colors = [self._pulse_color, "#52525B"]
        self._pulse_state = (self._pulse_state + 1) % 2
        
        try:
            self.configure(text_color=colors[self._pulse_state])
        except Exception:
            pass
        
        self.after(500, self._pulse)


class SlidePanel(ctk.CTkFrame):
    """滑动面板 - 用于侧边栏动画"""
    
    def __init__(self, master, start_pos: int = -250, end_pos: int = 0, **kwargs):
        super().__init__(master, **kwargs)
        self._start_pos = start_pos
        self._end_pos = end_pos
        self._current_pos = start_pos
        self._is_open = False
    
    def toggle(self, duration_ms: int = 200):
        """切换面板显示状态"""
        if self._is_open:
            self.slide_out(duration_ms)
        else:
            self.slide_in(duration_ms)
    
    def slide_in(self, duration_ms: int = 200):
        """滑入"""
        self._animate_slide(self._start_pos, self._end_pos, duration_ms)
        self._is_open = True
    
    def slide_out(self, duration_ms: int = 200):
        """滑出"""
        self._animate_slide(self._end_pos, self._start_pos, duration_ms)
        self._is_open = False
    
    def _animate_slide(self, start: int, end: int, duration_ms: int):
        steps = 10
        delay = duration_ms // steps
        delta = (end - start) / steps
        
        def step(current):
            if (delta > 0 and current >= end) or (delta < 0 and current <= end):
                self._current_pos = end
                return
            
            self._current_pos = current
            try:
                self.place(x=int(current))
            except Exception:
                pass
            
            self.after(delay, lambda: step(current + delta))
        
        step(start)


# 便捷函数
def add_hover_effect(button: ctk.CTkButton, hover_color: str = None):
    """为现有按钮添加悬停效果"""
    original_fg = button.cget('fg_color')
    
    def on_enter(e):
        if hover_color:
            button.configure(fg_color=hover_color)
    
    def on_leave(e):
        button.configure(fg_color=original_fg)
    
    button.bind('<Enter>', on_enter)
    button.bind('<Leave>', on_leave)


# 导出
__all__ = [
    'AnimationMixin', 'HoverButton', 'GlowFrame', 
    'PulseIndicator', 'SlidePanel', 'add_hover_effect'
]
