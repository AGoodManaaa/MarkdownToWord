# -*- coding: utf-8 -*-

"""与导出 Word 相关的 App 帮助函数"""

import os
import threading
from tkinter import filedialog, messagebox

import customtkinter as ctk

from converter import MarkdownToWordConverter
from ui.theme import COLORS


def export_to_word_for_app(app) -> None:
    """从 App 导出为 Word 文档的入口。

    保持原有行为：如果内容为空则提示，否则直接按默认样式导出。
    """
    content = app.input_text.get("1.0", "end-1c")
    if not content.strip():
        messagebox.showwarning("提示", "请先输入Markdown内容")
        return

    # 直接导出，使用默认设置
    do_export_for_app(app, content, "standard", "a4")


def show_export_options_for_app(app, content: str) -> None:
    """显示导出选项对话框。"""
    dialog = ctk.CTkToplevel(app)
    dialog.title("导出选项")
    dialog.geometry("400x350")
    dialog.transient(app)
    dialog.grab_set()

    # 居中显示
    dialog.update_idletasks()
    x = app.winfo_x() + (app.winfo_width() - 400) // 2
    y = app.winfo_y() + (app.winfo_height() - 350) // 2
    dialog.geometry(f"+{x}+{y}")

    # 标题
    ctk.CTkLabel(
        dialog,
        text="📄 导出设置",
        font=ctk.CTkFont(size=18, weight="bold"),
    ).pack(pady=(20, 15))

    # 样式选择
    style_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    style_frame.pack(fill="x", padx=30, pady=10)

    ctk.CTkLabel(
        style_frame,
        text="文档样式：",
        font=ctk.CTkFont(size=14),
    ).pack(anchor="w")

    style_var = ctk.StringVar(value="standard")

    styles = [
        ("standard", "📘 标准样式 - 宋体/Times New Roman"),
        ("academic", "🎓 学术论文 - 严格的学术格式"),
        ("simple", "✨ 简洁样式 - 干净简约"),
    ]

    for value, label in styles:
        ctk.CTkRadioButton(
            style_frame,
            text=label,
            variable=style_var,
            value=value,
            font=ctk.CTkFont(size=12),
        ).pack(anchor="w", pady=5, padx=10)

    # 页面设置
    page_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    page_frame.pack(fill="x", padx=30, pady=10)

    ctk.CTkLabel(
        page_frame,
        text="页面设置：",
        font=ctk.CTkFont(size=14),
    ).pack(anchor="w")

    page_var = ctk.StringVar(value="a4")
    page_options = ctk.CTkFrame(page_frame, fg_color="transparent")
    page_options.pack(fill="x", pady=5, padx=10)

    ctk.CTkRadioButton(page_options, text="A4", variable=page_var, value="a4").pack(
        side="left", padx=10
    )
    ctk.CTkRadioButton(
        page_options, text="Letter", variable=page_var, value="letter"
    ).pack(side="left", padx=10)

    # 按钮
    btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    btn_frame.pack(fill="x", padx=30, pady=20)

    def do_export() -> None:
        dialog.destroy()
        do_export_for_app(app, content, style_var.get(), page_var.get())

    def open_style_settings() -> None:
        try:
            if hasattr(app, 'open_export_style_settings'):
                app.open_export_style_settings()
        except Exception:
            pass

    ctk.CTkButton(
        btn_frame,
        text="📤 导出",
        command=do_export,
        fg_color=COLORS["primary"],
        width=120,
    ).pack(side="right", padx=5)

    ctk.CTkButton(
        btn_frame,
        text="⚙ 导出样式",
        command=open_style_settings,
        fg_color="transparent",
        border_width=1,
        border_color=COLORS["border"],
        text_color=COLORS["text_primary"],
        width=110,
    ).pack(side="left", padx=5)

    ctk.CTkButton(
        btn_frame,
        text="取消",
        command=dialog.destroy,
        fg_color="transparent",
        border_width=1,
        border_color=COLORS["border"],
        text_color=COLORS["text_primary"],
        width=80,
    ).pack(side="right", padx=5)


def do_export_for_app(app, content: str, style: str, page_size: str) -> None:
    """执行导出逻辑。"""
    # 选择保存路径
    default_name = (
        os.path.splitext(os.path.basename(app.current_file))[0]
        if app.current_file
        else "output"
    )
    file_path = filedialog.asksaveasfilename(
        title="保存Word文档",
        defaultextension=".docx",
        initialfile=f"{default_name}.docx",
        filetypes=[("Word文档", "*.docx")],
    )

    if not file_path:
        return

    app.update_status("⏳ 正在转换...")
    app.export_btn.configure(state="disabled")

    # 在线程中执行转换
    def convert() -> None:
        try:
            base_dir = (
                os.path.dirname(app.current_file) if app.current_file else os.getcwd()
            )
            converter = MarkdownToWordConverter(
                base_dir=base_dir,
                style=style,
                page_size=page_size,
                export_style=(app.config.get('export_style') if hasattr(app, 'config') else None),
            )
            converter.convert_text(content)
            converter.save(file_path)

            app.after(0, lambda fp=file_path: on_export_success_for_app(app, fp))
        except Exception as e:  # noqa: BLE001 - 保持原始广泛捕获
            error_msg = str(e)
            app.after(
                0, lambda msg=error_msg: on_export_error_for_app(app, msg)
            )

    threading.Thread(target=convert, daemon=True).start()


def on_export_success_for_app(app, file_path: str) -> None:
    """导出成功回调。"""
    app.export_btn.configure(state="normal")
    app.update_status(f"✅ 导出成功: {os.path.basename(file_path)}")

    if messagebox.askyesno(
        "导出成功", f"文档已保存到:\n{file_path}\n\n是否打开文件？"
    ):
        app._open_file_cross_platform(file_path)


def on_export_error_for_app(app, error: str) -> None:
    """导出失败回调。"""
    app.export_btn.configure(state="normal")
    app.update_status("❌ 导出失败")
    messagebox.showerror("导出错误", f"转换失败:\n{error}")
