# -*- coding: utf-8 -*-

"""与导出 Word 相关的 App 帮助函数"""

import os
import threading
import traceback
from threading import Event
from tkinter import filedialog, messagebox
import tkinter as tk

import customtkinter as ctk

from converter import MarkdownToWordConverter, ExportCancelled
from ui.preflight import run_preflight
from ui.export_history import record_export_event
from ui.theme import COLORS


def export_to_word_for_app(app) -> None:
    """从 App 导出为 Word 文档的入口。

    保持原有行为：如果内容为空则提示，否则直接按默认样式导出。
    """
    content = app.input_text.get("1.0", "end-1c")
    if not content.strip():
        messagebox.showwarning("提示", "请先输入Markdown内容")
        return

    # 优化：默认显示导出选项，并记住上次选择
    show_export_options_for_app(app, content)


def show_export_options_for_app(app, content: str) -> None:
    """显示导出选项对话框。"""
    dialog = ctk.CTkToplevel(app)
    dialog.title("导出选项")
    w, h = 400, 420
    dialog.geometry(f"{w}x{h}")
    dialog.transient(app)
    dialog.grab_set()

    # 居中显示
    dialog.update_idletasks()
    x = app.winfo_x() + (app.winfo_width() - w) // 2
    y = app.winfo_y() + (app.winfo_height() - h) // 2
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

    last_style = None
    last_page = None
    last_preflight_remote = None
    try:
        last_style = (app.config or {}).get('last_export_style')
        last_page = (app.config or {}).get('last_export_page_size')
        last_preflight_remote = (app.config or {}).get('preflight_check_remote_images')
    except Exception:
        last_style = None
        last_page = None
        last_preflight_remote = None

    style_var = ctk.StringVar(value=(last_style or "standard"))

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

    page_var = ctk.StringVar(value=(last_page or "a4"))
    page_options = ctk.CTkFrame(page_frame, fg_color="transparent")
    page_options.pack(fill="x", pady=5, padx=10)

    ctk.CTkRadioButton(page_options, text="A4", variable=page_var, value="a4").pack(
        side="left", padx=10
    )
    ctk.CTkRadioButton(
        page_options, text="Letter", variable=page_var, value="letter"
    ).pack(side="left", padx=10)

    # 预检查选项
    preflight_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    preflight_frame.pack(fill="x", padx=30, pady=(6, 0))

    ctk.CTkLabel(
        preflight_frame,
        text="导出前检查：",
        font=ctk.CTkFont(size=14),
    ).pack(anchor="w")

    remote_var = ctk.BooleanVar(value=bool(last_preflight_remote))
    ctk.CTkSwitch(
        preflight_frame,
        text="检查网络图片可访问性（可能稍慢）",
        variable=remote_var,
    ).pack(anchor="w", pady=(6, 0), padx=10)

    # 按钮
    btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    btn_frame.pack(fill="x", padx=30, pady=20)

    def do_export() -> None:
        dialog.destroy()
        try:
            if hasattr(app, 'config') and isinstance(app.config, dict):
                app.config['last_export_style'] = style_var.get()
                app.config['last_export_page_size'] = page_var.get()
                app.config['preflight_check_remote_images'] = bool(remote_var.get())
                try:
                    from ui.theme import save_config
                    save_config(app.config)
                except Exception:
                    pass
        except Exception:
            pass
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
    base_dir = (
        os.path.dirname(app.current_file) if getattr(app, 'current_file', None) else os.getcwd()
    )

    issues = []
    try:
        check_remote = False
        try:
            check_remote = bool((app.config or {}).get('preflight_check_remote_images', False))
        except Exception:
            check_remote = False

        issues = run_preflight(
            content,
            base_dir=base_dir,
            options={
                'check_remote_images': check_remote,
            },
        )
    except Exception:
        issues = []

    if issues:
        preview_lines = []
        for it in issues[:20]:
            ln = it.get('line')
            msg = it.get('message')
            hint = it.get('hint')
            prefix = f"L{ln}: " if ln else ""
            line = f"{prefix}{msg}" if msg else prefix
            if hint:
                line += f"\n  - {hint}"
            preview_lines.append(line)

        detail = "\n\n".join(preview_lines)
        if len(issues) > 20:
            detail += f"\n\n... 还有 {len(issues) - 20} 条未显示"

        if not messagebox.askyesno(
            "导出前检查",
            "检测到可能导致导出失败的问题：\n\n" + detail + "\n\n是否仍要继续导出？",
        ):
            return

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

    try:
        app._last_export_style = style
        app._last_export_page_size = page_size
        app._last_export_output_path = file_path
    except Exception:
        pass

    app.update_status("⏳ 正在转换...")
    app.export_btn.configure(state="disabled")
    try:
        if hasattr(app, 'cancel_export_btn'):
            app.cancel_export_btn.configure(state="normal")
    except Exception:
        pass

    cancel_event = Event()
    try:
        app._export_cancel_event = cancel_event
    except Exception:
        pass

    def on_progress(done: int, total: int, block_type: str, start_line: int) -> None:
        try:
            app.update_status(f"⏳ 正在转换... {done}/{total} ({block_type}, 行{start_line})")
        except Exception:
            pass

    # 在线程中执行转换
    def convert() -> None:
        try:
            converter = MarkdownToWordConverter(
                base_dir=base_dir,
                style=style,
                page_size=page_size,
                export_style=(app.config.get('export_style') if hasattr(app, 'config') else None),
            )
            converter.convert_text(content, progress_callback=on_progress, cancel_event=cancel_event)
            converter.save(file_path)

            app.after(0, lambda fp=file_path: on_export_success_for_app(app, fp))
        except ExportCancelled:
            app.after(0, lambda: on_export_cancel_for_app(app))
        except Exception as e:  # noqa: BLE001 - 保持原始广泛捕获
            tb = traceback.format_exc()
            error_msg = str(e)
            app.after(0, lambda msg=f"{error_msg}\n\n{tb}": on_export_error_for_app(app, msg))

    threading.Thread(target=convert, daemon=True).start()


def on_export_success_for_app(app, file_path: str) -> None:
    """导出成功回调。"""
    app.export_btn.configure(state="normal")
    try:
        if hasattr(app, 'cancel_export_btn'):
            app.cancel_export_btn.configure(state="disabled")
    except Exception:
        pass
    try:
        app._export_cancel_event = None
    except Exception:
        pass
    try:
        record_export_event(
            app,
            status='success',
            output_path=getattr(app, '_last_export_output_path', file_path),
            style=getattr(app, '_last_export_style', None),
            page_size=getattr(app, '_last_export_page_size', None),
        )
    except Exception:
        pass
    app.update_status(f"✅ 导出成功: {os.path.basename(file_path)}")

    if messagebox.askyesno(
        "导出成功", f"文档已保存到:\n{file_path}\n\n是否打开文件？"
    ):
        app._open_file_cross_platform(file_path)


def on_export_cancel_for_app(app) -> None:
    """导出取消回调。"""
    try:
        app.export_btn.configure(state="normal")
    except Exception:
        pass
    try:
        if hasattr(app, 'cancel_export_btn'):
            app.cancel_export_btn.configure(state="disabled")
    except Exception:
        pass
    try:
        app._export_cancel_event = None
    except Exception:
        pass
    try:
        record_export_event(
            app,
            status='cancelled',
            output_path=getattr(app, '_last_export_output_path', None),
            style=getattr(app, '_last_export_style', None),
            page_size=getattr(app, '_last_export_page_size', None),
        )
    except Exception:
        pass
    app.update_status("⛔ 已取消导出")


def on_export_error_for_app(app, error: str) -> None:
    """导出失败回调。"""
    app.export_btn.configure(state="normal")
    try:
        if hasattr(app, 'cancel_export_btn'):
            app.cancel_export_btn.configure(state="disabled")
    except Exception:
        pass
    try:
        app._export_cancel_event = None
    except Exception:
        pass
    app.update_status("❌ 导出失败")

    def _split_error(err: str) -> tuple[str, str]:
        s = str(err or "")
        s = s.replace("\r\n", "\n")
        if "\n\n" in s:
            summary = s.split("\n\n", 1)[0].strip()
            details = s.strip()
        else:
            lines = [ln for ln in s.split("\n") if ln.strip()]
            summary = (lines[0].strip() if lines else "导出失败")
            details = s.strip()
        if len(summary) > 400:
            summary = summary[:400] + "..."
        if len(details) > 20000:
            details = details[:20000] + "\n..."
        return summary, details

    def _show_error_dialog(summary: str, details: str) -> None:
        win = ctk.CTkToplevel(app)
        win.title("导出错误")
        win.geometry("720x460")
        win.transient(app)
        win.grab_set()

        container = ctk.CTkFrame(win, fg_color=COLORS['bg_card'])
        container.pack(fill='both', expand=True, padx=14, pady=14)

        ctk.CTkLabel(
            container,
            text="❌ 导出失败",
            font=ctk.CTkFont(size=18, weight='bold'),
            text_color=COLORS['text_primary'],
        ).pack(anchor='w', padx=12, pady=(10, 6))

        ctk.CTkLabel(
            container,
            text=summary,
            justify='left',
            wraplength=660,
            text_color=COLORS['text_primary'],
        ).pack(anchor='w', padx=12, pady=(0, 8))

        btns = ctk.CTkFrame(container, fg_color='transparent')
        btns.pack(fill='x', padx=12, pady=(0, 10))

        detail_frame = ctk.CTkFrame(container, fg_color='transparent')
        detail_visible = {'v': False}

        txt = tk.Text(detail_frame, height=12, wrap='word')
        txt.insert('1.0', details)
        txt.configure(state='disabled')
        txt.pack(fill='both', expand=True)

        def toggle_detail() -> None:
            if detail_visible['v']:
                try:
                    detail_frame.pack_forget()
                except Exception:
                    pass
                detail_visible['v'] = False
                try:
                    toggle_btn.configure(text='展开详情')
                except Exception:
                    pass
            else:
                detail_frame.pack(fill='both', expand=True, padx=12, pady=(0, 10))
                detail_visible['v'] = True
                try:
                    toggle_btn.configure(text='收起详情')
                except Exception:
                    pass

        def copy_detail() -> None:
            try:
                app.clipboard_clear()
                app.clipboard_append(details)
            except Exception:
                pass

        toggle_btn = ctk.CTkButton(
            btns,
            text='展开详情',
            fg_color='transparent',
            border_width=1,
            border_color=COLORS['border'],
            text_color=COLORS['text_primary'],
            command=toggle_detail,
            width=110,
        )
        toggle_btn.pack(side='left')

        ctk.CTkButton(
            btns,
            text='复制详情',
            fg_color='transparent',
            border_width=1,
            border_color=COLORS['border'],
            text_color=COLORS['text_primary'],
            command=copy_detail,
            width=110,
        ).pack(side='left', padx=8)

        ctk.CTkButton(
            btns,
            text='关闭',
            fg_color=COLORS['primary'],
            command=win.destroy,
            width=90,
        ).pack(side='right')

    summary, details = _split_error(error)
    try:
        record_export_event(
            app,
            status='error',
            output_path=getattr(app, '_last_export_output_path', None),
            style=getattr(app, '_last_export_style', None),
            page_size=getattr(app, '_last_export_page_size', None),
            error=str(summary)[:500],
        )
    except Exception:
        pass

    try:
        _show_error_dialog(summary, details)
    except Exception:
        messagebox.showerror("导出错误", f"转换失败:\n{summary}")
