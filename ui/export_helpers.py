# -*- coding: utf-8 -*-

"""与导出 Word 相关的 App 帮助函数"""

import os
import threading
import traceback
from threading import Event
from tkinter import filedialog, messagebox
import tkinter as tk
import time

import customtkinter as ctk

from converter import MarkdownToWordConverter, ExportCancelled
from ui.preflight import run_preflight
from ui.export_history import record_export_event
from ui.theme import COLORS, apply_window_icon, attach_window_geometry


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
    try:
        apply_window_icon(dialog)
    except Exception:
        pass
    try:
        attach_window_geometry(app, dialog, 'export_options')
    except Exception:
        pass
    w, h = 600, 750
    dialog.geometry(f"{w}x{h}")
    try:
        dialog.minsize(420, 520)
        dialog.resizable(True, True)
    except Exception:
        pass
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
        font=ctk.CTkFont(size=20, weight="bold"),
    ).pack(pady=(20, 15))

    # 样式选择
    style_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    style_frame.pack(fill="x", padx=30, pady=10)

    ctk.CTkLabel(
        style_frame,
        text="文档样式：",
        font=ctk.CTkFont(size=16),
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
            font=ctk.CTkFont(size=14),
        ).pack(anchor="w", pady=5, padx=10)

    # 模板选择
    template_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    template_frame.pack(fill="x", padx=30, pady=10)

    ctk.CTkLabel(
        template_frame,
        text="Word 模板：",
        font=ctk.CTkFont(size=16),
    ).pack(anchor="w")

    templates = ["默认样式"]
    if hasattr(app, 'template_manager'):
        templates = app.template_manager.list_templates()
        # 读取当前配置的模板名作为默认选项
        current_tpl = (app.template_manager.current_template or "默认样式")
    else:
        current_tpl = "默认样式"

    selected_template = ctk.StringVar(value=current_tpl if current_tpl in templates else "默认样式")
    template_combo = ctk.CTkComboBox(template_frame, values=templates, variable=selected_template, width=280)
    template_combo.pack(pady=5, padx=10, anchor="w", side="left")

    def import_template_cmd():
        if hasattr(app, 'template_manager'):
            filename = app.template_manager.quick_import()
            if filename:
                new_templates = app.template_manager.list_templates()
                template_combo.configure(values=new_templates)
                template_combo.set(filename)

    ctk.CTkButton(template_frame, text="导入", width=60, command=import_template_cmd).pack(pady=5, padx=5, anchor="w", side="left")

    # 页面设置
    page_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    page_frame.pack(fill="x", padx=30, pady=10)

    ctk.CTkLabel(
        page_frame,
        text="页面设置：",
        font=ctk.CTkFont(size=16),
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
        font=ctk.CTkFont(size=16),
    ).pack(anchor="w")

    remote_var = ctk.BooleanVar(value=bool(last_preflight_remote))
    ctk.CTkSwitch(
        preflight_frame,
        text="检查网络图片可访问性（可能稍慢）",
        variable=remote_var,
    ).pack(anchor="w", pady=(6, 0), padx=10)

    # 目录与字段更新
    toc_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    toc_frame.pack(fill="x", padx=30, pady=(10, 0))

    ctk.CTkLabel(
        toc_frame,
        text="目录（TOC）：",
        font=ctk.CTkFont(size=16),
    ).pack(anchor="w")

    toc_enabled = False
    update_fields = True
    try:
        toc_enabled = bool((app.config or {}).get('export_toc_enabled', False))
        update_fields = bool((app.config or {}).get('export_update_fields_on_open', True))
    except Exception:
        toc_enabled = False
        update_fields = True

    toc_var = ctk.BooleanVar(value=toc_enabled)
    update_fields_var = ctk.BooleanVar(value=update_fields)

    auto_format_enabled = False
    try:
        auto_format_enabled = bool((app.config or {}).get('export_auto_format_markdown', False))
    except Exception:
        auto_format_enabled = False
    auto_format_var = ctk.BooleanVar(value=auto_format_enabled)

    ctk.CTkSwitch(
        toc_frame,
        text="启用 [[TOC]] 标记生成目录",
        variable=toc_var,
    ).pack(anchor="w", pady=(6, 0), padx=10)

    ctk.CTkSwitch(
        toc_frame,
        text="打开文档时自动更新目录/编号",
        variable=update_fields_var,
    ).pack(anchor="w", pady=(6, 0), padx=10)

    ctk.CTkSwitch(
        toc_frame,
        text="导出前自动规范化 Markdown（推荐）",
        variable=auto_format_var,
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
                app.config['export_toc_enabled'] = bool(toc_var.get())
                app.config['export_update_fields_on_open'] = bool(update_fields_var.get())
                app.config['export_auto_format_markdown'] = bool(auto_format_var.get())
                try:
                    from ui.theme import save_config
                    save_config(app.config)
                except Exception:
                    pass
        except Exception:
            pass
            
        template_name = selected_template.get()
        template_path = None
        if hasattr(app, 'template_manager'):
            try:
                app.template_manager.select_template(template_name if template_name != "默认样式" else None)
            except Exception:
                pass
            template_path = app.template_manager.resolve_path(template_name)
            
        do_export_for_app(app, content, style_var.get(), page_var.get(), template_path)

    def do_export_pdf() -> None:
        dialog.destroy()
        try:
            if hasattr(app, 'config') and isinstance(app.config, dict):
                app.config['last_export_style'] = style_var.get()
                app.config['last_export_page_size'] = page_var.get()
                app.config['preflight_check_remote_images'] = bool(remote_var.get())
                app.config['export_toc_enabled'] = bool(toc_var.get())
                app.config['export_update_fields_on_open'] = bool(update_fields_var.get())
                app.config['export_auto_format_markdown'] = bool(auto_format_var.get())
                try:
                    from ui.theme import save_config
                    save_config(app.config)
                except Exception:
                    pass
        except Exception:
            pass
        # 调用 PDF 导出功能
        try:
            if hasattr(app, 'pdf_export_feature'):
                app.pdf_export_feature.export_to_pdf()
        except Exception:
            pass

    def open_style_settings() -> None:
        try:
            if hasattr(app, 'open_export_style_settings'):
                app.open_export_style_settings()
        except Exception:
            pass

    ctk.CTkButton(
        btn_frame,
        text="📤 导出 Word",
        command=do_export,
        fg_color=COLORS["primary"],
        width=110,
    ).pack(side="right", padx=5)

    ctk.CTkButton(
        btn_frame,
        text="📄 导出 PDF",
        command=do_export_pdf,
        fg_color=COLORS["success"],
        width=100,
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


def do_export_for_app(app, content: str, style: str, page_size: str, template_path: str = None) -> None:
    """执行导出逻辑。"""
    # 可选：导出前自动规范化 Markdown
    try:
        if bool((getattr(app, 'config', None) or {}).get('export_auto_format_markdown', False)):
            try:
                from ui.formatting import format_markdown_text
                content = format_markdown_text(content)
            except Exception:
                pass
    except Exception:
        pass

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

    try:
        app._last_preflight_issues = list(issues or [])
    except Exception:
        pass

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
    initial_dir = None
    try:
        last_path = (app.config or {}).get('last_export_output_path')
        if last_path:
            last_dir = os.path.dirname(str(last_path))
            if last_dir and os.path.isdir(last_dir):
                initial_dir = last_dir
    except Exception:
        initial_dir = None
    if not initial_dir:
        try:
            if app.current_file:
                initial_dir = os.path.dirname(app.current_file)
        except Exception:
            initial_dir = None

    initial_file = None
    try:
        last_path = (app.config or {}).get('last_export_output_path')
        if last_path:
            base = os.path.basename(str(last_path))
            if base.lower().endswith('.docx'):
                initial_file = base
    except Exception:
        initial_file = None
    if not initial_file:
        initial_file = f"{default_name}.docx"

    file_path = filedialog.asksaveasfilename(
        title="保存Word文档",
        defaultextension=".docx",
        initialfile=initial_file,
        initialdir=initial_dir,
        filetypes=[("Word文档", "*.docx")],
    )

    if not file_path:
        return

    # 进入忙碌态：禁用按钮，避免重复触发
    try:
        if hasattr(app, 'busy') and app.busy is not None:
            app.busy.enter('export', message='⏳ 正在转换...')
    except Exception:
        pass

    try:
        app._last_export_style = style
        app._last_export_page_size = page_size
        app._last_export_output_path = file_path
    except Exception:
        pass

    try:
        if hasattr(app, 'config') and isinstance(app.config, dict):
            app.config['last_export_output_path'] = file_path
            try:
                from ui.theme import save_config
                save_config(app.config)
            except Exception:
                pass
    except Exception:
        pass

    try:
        if hasattr(app, 'status_bar_feature') and app.status_bar_feature is not None:
            app.status_bar_feature.update_progress(0.0, "⏳ 正在转换...")
        else:
            app.update_status("⏳ 正在转换...")
    except Exception:
        app.update_status("⏳ 正在转换...")
    try:
        if hasattr(app, 'export_btn') and app.export_btn is not None:
            app.export_btn.configure(state="disabled")
    except Exception:
        pass
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

    last_ui_ts = 0.0
    last_pct = -1

    def on_progress(done: int, total: int, block_type: str, start_line: int) -> None:
        try:
            p = 0.0
            if total:
                p = float(done) / float(total)
            pct = int(p * 100)

            now = time.monotonic()
            # UI 刷新节流：避免进度回调过密导致界面抖动
            # 但确保最终 100% 或 done==total 时一定刷新
            min_interval = 0.12
            nonlocal last_ui_ts, last_pct
            if (done != total) and (pct == last_pct) and ((now - last_ui_ts) < min_interval):
                return
            if (done != total) and ((now - last_ui_ts) < min_interval):
                return
            last_ui_ts = now
            last_pct = pct

            msg = f"⏳ 正在转换... {pct}% ({done}/{total})  {block_type}  行{start_line}"

            def _apply() -> None:
                try:
                    if hasattr(app, 'status_bar_feature') and app.status_bar_feature is not None:
                        app.status_bar_feature.update_progress(p, msg)
                    else:
                        app.update_status(msg)
                except Exception:
                    pass

            try:
                app.after(0, _apply)
            except Exception:
                _apply()
        except Exception:
            pass

    # 在线程中执行转换
    def convert() -> None:
        try:
            export_style = None
            try:
                export_style = (app.config.get('export_style') if hasattr(app, 'config') else None)
                if not isinstance(export_style, dict):
                    export_style = {}
                export_style = {
                    **export_style,
                    'toc_enabled': bool((app.config or {}).get('export_toc_enabled', False)),
                    'update_fields_on_open': bool((app.config or {}).get('export_update_fields_on_open', True)),
                }
            except Exception:
                export_style = (app.config.get('export_style') if hasattr(app, 'config') else None)

            # 获取页眉页脚配置
            header_footer_config = None
            if hasattr(app, 'header_footer') and hasattr(app.header_footer, 'config'):
                header_footer_config = app.header_footer.config

            converter = MarkdownToWordConverter(
                base_dir=base_dir,
                style=style,
                page_size=page_size,
                export_style=export_style,
                template_path=template_path,
                header_footer_config=header_footer_config
            )
            converter.convert_text(content, progress_callback=on_progress, cancel_event=cancel_event)

            # 收集诊断信息（例如缺失图片）
            try:
                diag = {
                    'missing_images': [],
                    'converter_diagnostics': {},
                }
                try:
                    ih = getattr(converter, 'image_handler', None)
                    if ih is not None and hasattr(ih, 'get_issues'):
                        diag['missing_images'] = [it for it in (ih.get_issues() or []) if isinstance(it, dict)]
                except Exception:
                    pass
                try:
                    cd = getattr(converter, 'diagnostics', None)
                    if isinstance(cd, dict):
                        diag['converter_diagnostics'] = cd
                except Exception:
                    pass
                app._last_export_diagnostics = diag
            except Exception:
                pass

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
    try:
        if hasattr(app, 'export_btn') and app.export_btn is not None:
            app.export_btn.configure(state="normal")
    except Exception:
        pass
    try:
        if hasattr(app, 'status_bar_feature') and app.status_bar_feature is not None:
            app.status_bar_feature.update_progress(None, None)
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
            status='success',
            output_path=getattr(app, '_last_export_output_path', file_path),
            style=getattr(app, '_last_export_style', None),
            page_size=getattr(app, '_last_export_page_size', None),
        )
    except Exception:
        pass
    try:
        if hasattr(app, 'busy') and app.busy is not None:
            app.busy.exit()
    except Exception:
        pass
    app.update_status(f"✅ 导出成功: {os.path.basename(file_path)}")

    # 如果存在缺失图片等问题，提示保存诊断报告
    try:
        diag = getattr(app, '_last_export_diagnostics', None) or {}
        missing = diag.get('missing_images') if isinstance(diag, dict) else None
        if isinstance(missing, list) and missing:
            def _show_missing_dialog() -> None:
                win = ctk.CTkToplevel(app)
                win.title('导出完成（有告警）')
                try:
                    apply_window_icon(win)
                except Exception:
                    pass
                try:
                    attach_window_geometry(app, win, 'export_missing_images')
                except Exception:
                    pass
                win.geometry('720x520')
                win.transient(app)
                win.grab_set()

                try:
                    win.update_idletasks()
                    x = app.winfo_x() + (app.winfo_width() - 720) // 2
                    y = app.winfo_y() + (app.winfo_height() - 520) // 2
                    win.geometry(f'+{x}+{y}')
                except Exception:
                    pass

                container = ctk.CTkFrame(win, fg_color=COLORS['bg_card'])
                container.pack(fill='both', expand=True, padx=14, pady=14)

                ctk.CTkLabel(
                    container,
                    text='⚠ 导出完成（有告警）',
                    font=ctk.CTkFont(size=18, weight='bold'),
                    text_color=COLORS['text_primary'],
                ).pack(anchor='w', padx=12, pady=(10, 6))

                ctk.CTkLabel(
                    container,
                    text=f"检测到 {len(missing)} 个图片无法加载（可能是路径错误/文件缺失/网络图片不可用）。",
                    justify='left',
                    wraplength=660,
                    text_color=COLORS['text_primary'],
                ).pack(anchor='w', padx=12, pady=(0, 10))

                txt_frame = ctk.CTkFrame(container, fg_color='transparent')
                txt_frame.pack(fill='both', expand=True, padx=12, pady=(0, 10))

                detail_lines = []
                for it in missing[:200]:
                    try:
                        detail_lines.append(str(it))
                    except Exception:
                        pass
                if len(missing) > 200:
                    detail_lines.append(f'... 还有 {len(missing) - 200} 条未显示')

                txt = tk.Text(txt_frame, height=14, wrap='word')
                txt.insert('1.0', '\n'.join(detail_lines))
                txt.configure(state='disabled')
                txt.pack(fill='both', expand=True)

                btns = ctk.CTkFrame(container, fg_color='transparent')
                btns.pack(fill='x', padx=12, pady=(0, 10))

                def copy_list() -> None:
                    try:
                        app.clipboard_clear()
                        app.clipboard_append('\n'.join(detail_lines))
                        try:
                            app.update_status('✅ 已复制缺失图片列表')
                        except Exception:
                            pass
                    except Exception:
                        pass

                ctk.CTkButton(
                    btns,
                    text='复制列表',
                    fg_color='transparent',
                    border_width=1,
                    border_color=COLORS['border'],
                    text_color=COLORS['text_primary'],
                    command=copy_list,
                    width=110,
                ).pack(side='left')

                ctk.CTkButton(
                    btns,
                    text='保存诊断报告',
                    fg_color='transparent',
                    border_width=1,
                    border_color=COLORS['border'],
                    text_color=COLORS['text_primary'],
                    command=lambda: _save_diagnostic_report_for_app(app, error_details=None),
                    width=130,
                ).pack(side='left', padx=8)

                ctk.CTkButton(
                    btns,
                    text='关闭',
                    fg_color=COLORS['primary'],
                    command=win.destroy,
                    width=90,
                ).pack(side='right')

            try:
                _show_missing_dialog()
            except Exception:
                if messagebox.askyesno(
                    '导出完成（有告警）',
                    f"导出完成，但检测到 {len(missing)} 个图片无法加载。\n\n是否保存诊断报告？",
                ):
                    _save_diagnostic_report_for_app(app, error_details=None)
    except Exception:
        pass

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
        if hasattr(app, 'status_bar_feature') and app.status_bar_feature is not None:
            app.status_bar_feature.update_progress(None, None)
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
    try:
        if hasattr(app, 'busy') and app.busy is not None:
            app.busy.exit("⛔ 已取消导出")
    except Exception:
        pass
    app.update_status("⛔ 已取消导出")


def on_export_error_for_app(app, error: str) -> None:
    """导出失败回调。"""
    try:
        if hasattr(app, 'export_btn') and app.export_btn is not None:
            app.export_btn.configure(state="normal")
    except Exception:
        pass
    try:
        if hasattr(app, 'status_bar_feature') and app.status_bar_feature is not None:
            app.status_bar_feature.update_progress(None, None)
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
        if hasattr(app, 'busy') and app.busy is not None:
            app.busy.exit("❌ 导出失败")
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
        try:
            apply_window_icon(win)
        except Exception:
            pass
        try:
            attach_window_geometry(app, win, 'export_error')
        except Exception:
            pass
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

        def copy_summary() -> None:
            try:
                app.clipboard_clear()
                app.clipboard_append(summary)
                try:
                    if hasattr(app, 'update_status'):
                        app.update_status('✅ 已复制摘要')
                except Exception:
                    pass
            except Exception:
                pass

        def copy_detail() -> None:
            try:
                app.clipboard_clear()
                app.clipboard_append(details)
                try:
                    if hasattr(app, 'update_status'):
                        app.update_status('✅ 已复制详情')
                except Exception:
                    pass
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
            text='复制摘要',
            fg_color='transparent',
            border_width=1,
            border_color=COLORS['border'],
            text_color=COLORS['text_primary'],
            command=copy_summary,
            width=110,
        ).pack(side='left')

        ctk.CTkButton(
            btns,
            text='保存诊断报告',
            fg_color='transparent',
            border_width=1,
            border_color=COLORS['border'],
            text_color=COLORS['text_primary'],
            command=lambda d=details: _save_diagnostic_report_for_app(app, error_details=d),
            width=130,
        ).pack(side='left')

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


def _build_diagnostic_report_for_app(app, error_details: str = None) -> str:
    lines = []
    try:
        lines.append('MarkdownToWord 诊断报告')
        lines.append('')

        try:
            import platform
            lines.append(f"OS: {platform.platform()}")
        except Exception:
            pass
        try:
            import sys
            lines.append(f"Python: {sys.version}")
        except Exception:
            pass
        lines.append('')

        cfg = getattr(app, 'config', None)
        if isinstance(cfg, dict):
            lines.append('--- 配置（节选）---')
            try:
                lines.append(f"last_export_style: {cfg.get('last_export_style')}")
                lines.append(f"last_export_page_size: {cfg.get('last_export_page_size')}")
                lines.append(f"export_toc_enabled: {cfg.get('export_toc_enabled')}")
                lines.append(f"export_update_fields_on_open: {cfg.get('export_update_fields_on_open')}")
                lines.append(f"preflight_check_remote_images: {cfg.get('preflight_check_remote_images')}")
            except Exception:
                pass
            lines.append('')

        lines.append('--- 文件信息 ---')
        try:
            lines.append(f"current_file: {getattr(app, 'current_file', None)}")
        except Exception:
            pass
        try:
            lines.append(f"last_output_path: {getattr(app, '_last_export_output_path', None)}")
        except Exception:
            pass
        lines.append('')

        pf = getattr(app, '_last_preflight_issues', None)
        if isinstance(pf, list) and pf:
            lines.append('--- 导出前检查（preflight）---')
            for it in pf[:200]:
                try:
                    lines.append(str(it))
                except Exception:
                    pass
            lines.append('')

        diag = getattr(app, '_last_export_diagnostics', None)
        if isinstance(diag, dict):
            missing = diag.get('missing_images')
            if isinstance(missing, list) and missing:
                lines.append('--- 缺失图片 ---')
                for it in missing[:500]:
                    try:
                        lines.append(str(it))
                    except Exception:
                        pass
                lines.append('')

            cd = diag.get('converter_diagnostics')
            if isinstance(cd, dict) and cd:
                try:
                    inc = cd.get('include_issues')
                    if isinstance(inc, list) and inc:
                        lines.append('--- include 问题 ---')
                        for it in inc[:500]:
                            try:
                                lines.append(str(it))
                            except Exception:
                                pass
                        lines.append('')
                except Exception:
                    pass

                try:
                    inc_files = cd.get('included_files')
                    if isinstance(inc_files, list) and inc_files:
                        lines.append('--- include 文件列表 ---')
                        for it in inc_files[:500]:
                            try:
                                lines.append(str(it))
                            except Exception:
                                pass
                        lines.append('')
                except Exception:
                    pass

                try:
                    ur = cd.get('unresolved_refs')
                    if isinstance(ur, list) and ur:
                        lines.append('--- 未解析引用 ---')
                        # 去重保持顺序
                        seen = set()
                        for it in ur:
                            try:
                                s = str(it)
                            except Exception:
                                continue
                            if s in seen:
                                continue
                            seen.add(s)
                            lines.append(s)
                        lines.append('')
                except Exception:
                    pass

        if error_details:
            lines.append('--- 错误详情 ---')
            try:
                lines.append(str(error_details))
            except Exception:
                pass
            lines.append('')
    except Exception:
        pass

    return "\n".join(lines).strip() + "\n"


def _save_diagnostic_report_for_app(app, error_details: str = None) -> None:
    try:
        report = _build_diagnostic_report_for_app(app, error_details=error_details)
        default_name = 'md2word_diagnostic.txt'

        initial_dir = None
        try:
            initial_dir = (getattr(app, 'config', None) or {}).get('last_save_dir')
            if initial_dir and not os.path.isdir(str(initial_dir)):
                initial_dir = None
        except Exception:
            initial_dir = None

        path = filedialog.asksaveasfilename(
            title='保存诊断报告',
            defaultextension='.txt',
            initialfile=default_name,
            initialdir=initial_dir,
            filetypes=[('Text', '*.txt')],
        )
        if not path:
            return
        with open(path, 'w', encoding='utf-8') as f:
            f.write(report)

        try:
            cfg = getattr(app, 'config', None)
            if isinstance(cfg, dict):
                cfg['last_save_dir'] = os.path.dirname(path)
                try:
                    from ui.theme import save_config
                    save_config(cfg)
                except Exception:
                    pass
        except Exception:
            pass

        try:
            if hasattr(app, 'update_status'):
                app.update_status(f"✅ 已保存诊断报告: {os.path.basename(path)}")
        except Exception:
            pass
    except Exception:
        pass
