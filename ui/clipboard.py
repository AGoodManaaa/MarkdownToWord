# -*- coding: utf-8 -*-

"""与剪贴板复制相关的 App 帮助函数"""

import os
import tempfile
from tkinter import messagebox

import markdown as md_parser

from converter import MarkdownToWordConverter


def copy_to_clipboard_for_app(app) -> None:
    """复制当前内容到剪贴板（Word 兼容）。"""
    content = app.input_text.get("1.0", "end-1c")
    if not content.strip():
        messagebox.showwarning("提示", "请先输入Markdown内容")
        return

    app.update_status("⏳ 正在生成剪贴板内容...")

    def copy_task() -> None:
        try:
            # 生成临时 Word 文档
            temp_file = tempfile.mktemp(suffix=".docx")
            base_dir = (
                os.path.dirname(app.current_file) if app.current_file else os.getcwd()
            )
            converter = MarkdownToWordConverter(base_dir=base_dir)
            converter.convert_text(content)
            converter.save(temp_file)

            # 使用 COM 复制到剪贴板
            copy_word_to_clipboard_for_app(app, temp_file)

            # 清理临时文件
            os.remove(temp_file)

            app.after(
                0,
                lambda: app.update_status(
                    "✅ 已复制到剪贴板，可直接粘贴到Word"
                ),
            )
            app.after(0, lambda: show_copy_toast_for_app(app))

        except Exception as e:  # noqa: BLE001 - 保持原始广泛捕获
            app.after(0, lambda: app.update_status(f"❌ 复制失败: {e}"))
            app.after(
                0,
                lambda: messagebox.showerror("错误", f"复制失败:\n{e}"),
            )

    # 与 GUI 线程分离
    import threading

    threading.Thread(target=copy_task, daemon=True).start()


def copy_word_to_clipboard_for_app(app, docx_path: str) -> None:
    """使用 COM 将 Word 内容复制到剪贴板。"""
    try:
        import win32com.client
        import pythoncom

        pythoncom.CoInitialize()

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        doc = word.Documents.Open(os.path.abspath(docx_path))
        doc.Content.Copy()  # 复制全部内容（包括公式）
        doc.Close(False)
        word.Quit()

        pythoncom.CoUninitialize()

    except ImportError:
        # 如果没有 pywin32，使用 HTML 格式复制
        copy_as_html_for_app(app, docx_path)


def copy_as_html_for_app(app, docx_path: str) -> None:  # noqa: ARG001 - 保留签名
    """备用方案：将 Markdown 转为 HTML 粘贴。"""
    content = app.input_text.get("1.0", "end-1c")

    # 转换为 HTML
    md = md_parser.Markdown(extensions=["tables", "fenced_code"])
    html = md.convert(content)

    # 添加简单样式
    styled_html = f"""
    <html>
    <head><meta charset="utf-8"></head>
    <body style="font-family: 'Times New Roman', serif; font-size: 11pt;">
    {html}
    </body>
    </html>
    """

    app.clipboard_clear()
    app.clipboard_append(styled_html)


def show_copy_toast_for_app(app) -> None:
    """显示复制成功的提示气泡。"""
    import customtkinter as ctk
    from ui.theme import COLORS

    toast = ctk.CTkToplevel(app)
    toast.overrideredirect(True)
    toast.attributes("-topmost", True)

    # 计算位置
    x = app.winfo_x() + app.winfo_width() // 2 - 150
    y = app.winfo_y() + app.winfo_height() - 100
    toast.geometry(f"300x50+{x}+{y}")

    frame = ctk.CTkFrame(toast, fg_color=COLORS["success"], corner_radius=12)
    frame.pack(fill="both", expand=True, padx=2, pady=2)

    label = ctk.CTkLabel(
        frame,
        text="✅ 已复制！可直接粘贴到Word",
        font=ctk.CTkFont(size=14, weight="bold"),
        text_color="white",
    )
    label.pack(expand=True)

    # 2 秒后自动关闭
    toast.after(2000, toast.destroy)
