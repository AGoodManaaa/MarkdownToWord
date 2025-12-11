"""Hotkey configuration dialog."""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Callable

import os
from ...utils.logging import log
from ...utils.win32 import HotkeyChecker, get_dpi_scale
from ...utils.resources import resource_path
from ...domains.hotkey.recorder import HotkeyRecorder
from ...i18n import t
from ...core.state import app_state


class HotkeyDialog:
    """热键设置对话框"""
    
    def __init__(self, current_hotkey: str, on_save: Callable[[str], None], on_close: Optional[Callable[[], None]] = None):
        """
        初始化热键设置对话框
        
        Args:
            current_hotkey: 当前热键字符串
            on_save: 保存回调函数，接收新的热键字符串
            on_close: 关闭回调函数（用于恢复热键监听）
        """
        self.current_hotkey = current_hotkey
        self.on_save = on_save
        self.on_close_callback = on_close
        self.new_hotkey: Optional[str] = None
        
        if app_state.root:
            self.root = tk.Toplevel(app_state.root)
        else:
            self.root = tk.Tk()
            
        self.root.title(t("hotkey.dialog.title"))
        
        # 设置图标
        try:
            icon_path = resource_path("assets/icons/logo.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception as e:
            log(f"Failed to set hotkey dialog icon: {e}")
        
        # 适配高分屏：根据 DPI 缩放窗口大小
        scale = get_dpi_scale()
        width = int(450 * scale)
        height = int(300 * scale)
        self.root.geometry(f"{width}x{height}")
        
        self.root.resizable(False, False)
        
        # 设置关闭窗口时的处理
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        
        # 窗口居中
        self._center_window()
        
        # 创建UI组件
        self._create_widgets()
        
        # 热键录制器
        self.recorder = HotkeyRecorder()
    
    def is_alive(self) -> bool:
        """判断窗口是否仍然存在"""
        try:
            return bool(self.root.winfo_exists())
        except Exception:
            return False
    
    def restore_and_focus(self):
        """从最小化恢复并置于前台"""
        if not self.is_alive():
            return
        
        try:
            self.root.deiconify()
            original_topmost = bool(self.root.attributes("-topmost"))
            self.root.lift()
            # 通过临时置顶确保窗口浮到前面，再恢复原状态
            self.root.attributes("-topmost", True)
            self.root.after(50, lambda: self.root.attributes("-topmost", original_topmost))
            self.root.focus_force()
        except Exception as e:
            log(f"Failed to restore hotkey dialog: {e}")
    
    def _center_window(self):
        """将窗口居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def _create_widgets(self):
        """创建UI组件"""
        # 标题
        title_frame = ttk.Frame(self.root, padding="10")
        title_frame.pack(fill=tk.X)
        
        title_label = ttk.Label(
            title_frame,
            text=t("hotkey.dialog.heading"),
            font=("Microsoft YaHei UI", 12, "bold")
        )
        title_label.pack()
        
        # 当前热键显示
        current_frame = ttk.Frame(self.root, padding="10")
        current_frame.pack(fill=tk.X)
        
        ttk.Label(current_frame, text=t("hotkey.dialog.current_hotkey")).pack(side=tk.LEFT)
        ttk.Label(
            current_frame,
            text=self._format_hotkey(self.current_hotkey),
            font=("Consolas", 10, "bold")
        ).pack(side=tk.LEFT, padx=5)
        
        # 新热键输入
        input_frame = ttk.Frame(self.root, padding="10")
        input_frame.pack(fill=tk.X)
        
        ttk.Label(input_frame, text=t("hotkey.dialog.new_hotkey")).pack(side=tk.LEFT)
        
        self.hotkey_entry = ttk.Entry(
            input_frame,
            font=("Consolas", 10),
            state="readonly",
            width=25
        )
        self.hotkey_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 录制按钮
        record_frame = ttk.Frame(self.root, padding="10")
        record_frame.pack(fill=tk.X)
        
        self.record_btn = ttk.Button(
            record_frame,
            text=t("hotkey.dialog.record_button"),
            command=self._start_recording
        )
        self.record_btn.pack()
        
        # 按钮栏
        button_frame = ttk.Frame(self.root, padding="10")
        button_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        cancel_btn = ttk.Button(
            button_frame,
            text=t("hotkey.dialog.cancel_button"),
            command=self._on_cancel,
            width=12
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)
        
        self.save_btn = ttk.Button(
            button_frame,
            text=t("hotkey.dialog.save_button"),
            command=self._on_save,
            state=tk.DISABLED,
            width=12
        )
        self.save_btn.pack(side=tk.RIGHT, padx=5)
    
    def _format_hotkey(self, hotkey: str) -> str:
        """格式化热键显示"""
        return hotkey.replace("<", "").replace(">", "").replace("+", " + ").title()
    
    def _start_recording(self):
        """开始录制热键"""
        self.new_hotkey = None
        
        self.record_btn.config(text=t("hotkey.dialog.recording_button"), state=tk.DISABLED)
        self.hotkey_entry.config(state=tk.NORMAL)
        self.hotkey_entry.delete(0, tk.END)
        self.hotkey_entry.insert(0, t("hotkey.dialog.waiting_input"))
        self.hotkey_entry.config(state="readonly")
        
        # 使用录制器开始录制
        self.recorder.start_recording(
            on_update=self._on_recording_update,
            on_finish=self._on_recording_finish
        )
    
    def _on_recording_update(self, display_text: str):
        """录制更新回调（实时显示）"""
        self.root.after(0, lambda: self._set_entry_text(display_text))
    
    def _on_recording_finish(self, hotkey_str: Optional[str], error: Optional[str]):
        """录制完成回调"""
        if error:
            # 有错误，显示警告
            self.root.after(0, lambda: messagebox.showwarning(t("hotkey.dialog.invalid_title"), error))
            self.root.after(0, self._reset_recording)
        elif hotkey_str:
            # 成功录制
            self.new_hotkey = hotkey_str
            self.root.after(0, self._enable_save_button)
        else:
            # 未知情况
            self.root.after(0, self._reset_recording)
    
    def _set_entry_text(self, text: str):
        """设置输入框文本（线程安全）"""
        self.hotkey_entry.config(state=tk.NORMAL)
        self.hotkey_entry.delete(0, tk.END)
        self.hotkey_entry.insert(0, text)
        self.hotkey_entry.config(state="readonly")
    
    def _enable_save_button(self):
        """启用保存按钮"""
        self.save_btn.config(state=tk.NORMAL)
        self.record_btn.config(text=t("hotkey.dialog.record_again"), state=tk.NORMAL)
    
    def _reset_recording(self):
        """重置录制状态"""
        self.record_btn.config(text=t("hotkey.dialog.record_button"), state=tk.NORMAL)
        self.hotkey_entry.config(state=tk.NORMAL)
        self.hotkey_entry.delete(0, tk.END)
        self.hotkey_entry.config(state="readonly")
    
    def _on_save(self):
        """保存热键"""
        if not self.new_hotkey:
            messagebox.showwarning(t("hotkey.dialog.notice_title"), t("hotkey.dialog.record_first"))
            return
        
        # 确认对话框
        confirm_msg = t(
            "hotkey.dialog.confirm_message",
            old_hotkey=self._format_hotkey(self.current_hotkey),
            new_hotkey=self._format_hotkey(self.new_hotkey)
        )
        
        if not messagebox.askyesno(t("hotkey.dialog.confirm_title"), confirm_msg):
            return
        
        # 检测热键占用
        if not HotkeyChecker.is_hotkey_available(self.new_hotkey):
            conflict_msg = t(
                "hotkey.dialog.conflict_message",
                hotkey=self._format_hotkey(self.new_hotkey)
            )
            if not messagebox.askyesno(t("hotkey.dialog.conflict_title"), conflict_msg, icon='warning'):
                return

        try:
            # 调用保存回调
            self.on_save(self.new_hotkey)
            messagebox.showinfo("成功", f"热键已更新为：{self._format_hotkey(self.new_hotkey)}\n\n请使用新热键测试功能。")
            self._cleanup()
            self._safe_destroy()
        except Exception as e:
            log(f"Failed to save hotkey: {e}")
            messagebox.showerror("错误", f"保存热键失败：{str(e)}")
    
    def _cleanup(self):
        """清理资源"""
        try:
            self.recorder.stop_recording()
        except Exception as e:
            log(f"Error stopping recorder: {e}")
    
    def _safe_destroy(self):
        """安全销毁窗口（避免线程问题）"""
        try:
            self.root.destroy()
        except Exception as e:
            # 忽略 Tcl_AsyncDelete 错误，这是 tkinter 在非主线程中的已知问题
            if "Tcl_AsyncDelete" not in str(e):
                log(f"Error destroying window: {e}")
    
    def _on_close(self):
        """窗口关闭事件"""
        self._cleanup()
        if self.on_close_callback:
            try:
                self.on_close_callback()
            except Exception as e:
                log(f"Error in close callback: {e}")
        self._safe_destroy()
    
    def _on_cancel(self):
        """取消设置"""
        self._cleanup()
        if self.on_close_callback:
            try:
                self.on_close_callback()
            except Exception as e:
                log(f"Error in close callback: {e}")
        self._safe_destroy()
    
    def show(self):
        """显示对话框"""
        try:
            self.root.transient(app_state.root)
            self.root.deiconify()
            self.restore_and_focus()
        except Exception as e:
            log(f"Error showing hotkey dialog: {e}")
