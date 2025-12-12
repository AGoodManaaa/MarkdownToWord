# -*- coding: utf-8 -*-

import customtkinter as ctk
from tkinter import filedialog, messagebox

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn

from ui.theme import COLORS, save_config, get_default_export_style


class ModernButton(ctk.CTkButton):
    """现代化按钮组件"""
    def __init__(self, master, text, command=None, style="primary", icon=None, **kwargs):
        colors = {
            'primary': (COLORS['primary'], COLORS['primary_hover']),
            'secondary': (COLORS['secondary'], '#E11D48'),
            'success': (COLORS['success'], '#16A34A'),
            'outline': ('transparent', COLORS['highlight']),
            'ghost': ('transparent', COLORS['highlight']),
            'danger': (COLORS['danger'], '#DC2626'),
        }
        
        fg_color, hover_color = colors.get(style, colors['primary'])
        is_outline = style == 'outline'
        is_ghost = style == 'ghost'
        is_light = is_outline or is_ghost
        text_color = 'white' if not is_light else COLORS['text_primary']
        border_width = 1 if is_outline else 0
        
        super().__init__(
            master,
            text=text,
            command=command,
            fg_color=fg_color,
            hover_color=hover_color,
            text_color=text_color,
            border_width=border_width,
            border_color=COLORS['border'] if is_outline else None,
            corner_radius=14,
            height=38,
            font=ctk.CTkFont(size=13, weight="bold"),
            **kwargs
        )


class ModernCard(ctk.CTkFrame):
    """现代化卡片容器"""
    def __init__(self, master, title=None, **kwargs):
        super().__init__(
            master,
            fg_color=COLORS['bg_card'],
            corner_radius=18,
            border_width=1,
            border_color=COLORS['border'],
            **kwargs
        )
        
        if title:
            self.title_label = ctk.CTkLabel(
                self,
                text=title,
                font=ctk.CTkFont(size=15, weight="bold"),
                text_color=COLORS['text_primary']
            )
            self.title_label.pack(anchor="w", padx=20, pady=(15, 10))


class ExportStyleSettingsDialog(ctk.CTkToplevel):
    def __init__(self, app):
        super().__init__(app)
        self.app = app

        self.title("导出样式设置")
        self.geometry("720x720")
        self.transient(app)
        self.grab_set()

        try:
            self.update_idletasks()
            x = app.winfo_x() + (app.winfo_width() - 720) // 2
            y = app.winfo_y() + (app.winfo_height() - 720) // 2
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

        self._vars = {}
        self._field_types = {}

        base = get_default_export_style()
        cur = {}
        try:
            cur = app.config.get('export_style') if hasattr(app, 'config') else {}
        except Exception:
            cur = {}
        if not isinstance(cur, dict):
            cur = {}
        self._style = {**base, **cur}

        container = ctk.CTkFrame(self, fg_color=COLORS['bg_card'])
        container.pack(fill='both', expand=True, padx=14, pady=14)

        ctk.CTkLabel(
            container,
            text="⚙ 导出样式设置",
            font=ctk.CTkFont(size=18, weight='bold'),
            text_color=COLORS['text_primary'],
        ).pack(anchor='w', padx=12, pady=(10, 8))

        scroll = ctk.CTkScrollableFrame(container, fg_color='transparent')
        scroll.pack(fill='both', expand=True, padx=10, pady=(0, 10))

        def add_row(key: str, label: str, kind: str, options=None):
            row = ctk.CTkFrame(scroll, fg_color='transparent')
            row.pack(fill='x', pady=6)

            ctk.CTkLabel(
                row,
                text=label,
                width=240,
                anchor='w',
                text_color=COLORS['text_primary'],
            ).pack(side='left')

            self._field_types[key] = kind
            value = self._style.get(key)

            if kind == 'bool':
                var = ctk.BooleanVar(value=bool(value))
                self._vars[key] = var
                w = ctk.CTkSwitch(row, text="", variable=var)
                w.pack(side='right')
                return

            if kind == 'option':
                var = ctk.StringVar(value=str(value) if value is not None else '')
                self._vars[key] = var
                w = ctk.CTkOptionMenu(row, variable=var, values=list(options or []))
                w.pack(side='right', fill='x', expand=True)
                return

            var = ctk.StringVar(value='' if value is None else str(value))
            self._vars[key] = var
            w = ctk.CTkEntry(row, textvariable=var)
            w.pack(side='right', fill='x', expand=True)

        align_opts = ['left', 'center', 'right', 'justify']
        pos_opts = ['before', 'after']

        ctk.CTkLabel(scroll, text="字体", font=ctk.CTkFont(size=14, weight='bold'), text_color=COLORS['text_primary']).pack(anchor='w', pady=(6, 2))
        add_row('body_cn', '正文中文字体 body_cn', 'str')
        add_row('body_en', '正文英文字体 body_en', 'str')
        add_row('heading_cn', '标题中文字体 heading_cn', 'str')
        add_row('mono', '等宽字体 mono', 'str')
        add_row('math', '数学字体 math', 'str')
        add_row('quote_font', '引用字体 quote_font', 'str')
        add_row('caption_font', '图表标题字体 caption_font', 'str')

        ctk.CTkLabel(scroll, text="字号（pt）", font=ctk.CTkFont(size=14, weight='bold'), text_color=COLORS['text_primary']).pack(anchor='w', pady=(14, 2))
        add_row('body_size_pt', '正文 body_size_pt', 'float')
        add_row('heading1_size_pt', '标题1 heading1_size_pt', 'float')
        add_row('heading2_size_pt', '标题2 heading2_size_pt', 'float')
        add_row('heading3_size_pt', '标题3 heading3_size_pt', 'float')
        add_row('heading4_size_pt', '标题4 heading4_size_pt', 'float')
        add_row('code_size_pt', '代码 code_size_pt', 'float')
        add_row('quote_size_pt', '引用 quote_size_pt', 'float')
        add_row('caption_size_pt', '图表标题 caption_size_pt', 'float')
        add_row('hyperlink_size_pt', '链接 hyperlink_size_pt', 'float')

        ctk.CTkLabel(scroll, text="段落", font=ctk.CTkFont(size=14, weight='bold'), text_color=COLORS['text_primary']).pack(anchor='w', pady=(14, 2))
        add_row('body_alignment', '正文对齐 body_alignment', 'option', align_opts)
        add_row('body_line_spacing', '正文行距 body_line_spacing', 'float')
        add_row('body_space_before_pt', '正文段前 body_space_before_pt', 'float')
        add_row('body_space_after_pt', '正文段后 body_space_after_pt', 'float')
        add_row('body_first_line_indent_pt', '正文首行缩进 body_first_line_indent_pt', 'float')

        ctk.CTkLabel(scroll, text="标题段落", font=ctk.CTkFont(size=14, weight='bold'), text_color=COLORS['text_primary']).pack(anchor='w', pady=(14, 2))
        add_row('heading1_alignment', 'H1 对齐 heading1_alignment', 'option', align_opts)
        add_row('heading1_space_before_pt', 'H1 段前 heading1_space_before_pt', 'float')
        add_row('heading1_space_after_pt', 'H1 段后 heading1_space_after_pt', 'float')
        add_row('heading1_bold', 'H1 加粗 heading1_bold', 'bool')
        add_row('heading2_alignment', 'H2 对齐 heading2_alignment', 'option', align_opts)
        add_row('heading2_space_before_pt', 'H2 段前 heading2_space_before_pt', 'float')
        add_row('heading2_space_after_pt', 'H2 段后 heading2_space_after_pt', 'float')
        add_row('heading2_bold', 'H2 加粗 heading2_bold', 'bool')
        add_row('heading3_alignment', 'H3 对齐 heading3_alignment', 'option', align_opts)
        add_row('heading3_space_before_pt', 'H3 段前 heading3_space_before_pt', 'float')
        add_row('heading3_space_after_pt', 'H3 段后 heading3_space_after_pt', 'float')
        add_row('heading3_bold', 'H3 加粗 heading3_bold', 'bool')
        add_row('heading4_alignment', 'H4 对齐 heading4_alignment', 'option', align_opts)
        add_row('heading4_space_before_pt', 'H4 段前 heading4_space_before_pt', 'float')
        add_row('heading4_space_after_pt', 'H4 段后 heading4_space_after_pt', 'float')
        add_row('heading4_bold', 'H4 加粗 heading4_bold', 'bool')

        ctk.CTkLabel(scroll, text="页边距（cm）", font=ctk.CTkFont(size=14, weight='bold'), text_color=COLORS['text_primary']).pack(anchor='w', pady=(14, 2))
        add_row('margin_top_cm', '上 margin_top_cm', 'float')
        add_row('margin_bottom_cm', '下 margin_bottom_cm', 'float')
        add_row('margin_left_cm', '左 margin_left_cm', 'float')
        add_row('margin_right_cm', '右 margin_right_cm', 'float')

        ctk.CTkLabel(scroll, text="引用块", font=ctk.CTkFont(size=14, weight='bold'), text_color=COLORS['text_primary']).pack(anchor='w', pady=(14, 2))
        add_row('quote_italic', '引用斜体 quote_italic', 'bool')
        add_row('quote_left_indent_cm', '引用左缩进 quote_left_indent_cm', 'float')
        add_row('quote_right_indent_cm', '引用右缩进 quote_right_indent_cm', 'float')
        add_row('quote_space_before_pt', '引用段前 quote_space_before_pt', 'float')
        add_row('quote_space_after_pt', '引用段后 quote_space_after_pt', 'float')

        ctk.CTkLabel(scroll, text="代码块", font=ctk.CTkFont(size=14, weight='bold'), text_color=COLORS['text_primary']).pack(anchor='w', pady=(14, 2))
        add_row('code_space_before_pt', '代码段前 code_space_before_pt', 'float')
        add_row('code_space_after_pt', '代码段后 code_space_after_pt', 'float')
        add_row('code_left_indent_cm', '代码左缩进 code_left_indent_cm', 'float')
        add_row('code_line_spacing', '代码行距 code_line_spacing', 'float')

        ctk.CTkLabel(scroll, text="链接", font=ctk.CTkFont(size=14, weight='bold'), text_color=COLORS['text_primary']).pack(anchor='w', pady=(14, 2))
        add_row('hyperlink_color', '链接颜色 hyperlink_color (RRGGBB)', 'str')
        add_row('hyperlink_underline', '链接下划线 hyperlink_underline', 'bool')

        ctk.CTkLabel(scroll, text="图片", font=ctk.CTkFont(size=14, weight='bold'), text_color=COLORS['text_primary']).pack(anchor='w', pady=(14, 2))
        add_row('image_max_width_in', '图片最大宽度 image_max_width_in (inch)', 'float')
        add_row('image_caption_position', '图片标题位置 image_caption_position', 'option', pos_opts)
        add_row('image_caption_align', '图片标题对齐 image_caption_align', 'option', align_opts)
        add_row('image_caption_template', '图片标题模板 image_caption_template', 'str')

        ctk.CTkLabel(scroll, text="表格", font=ctk.CTkFont(size=14, weight='bold'), text_color=COLORS['text_primary']).pack(anchor='w', pady=(14, 2))
        add_row('table_three_line', '三线表 table_three_line', 'bool')
        add_row('table_alignment', '表格对齐 table_alignment', 'option', ['left', 'center', 'right'])
        add_row('table_header_bold', '表头加粗 table_header_bold', 'bool')
        add_row('table_caption_position', '表格标题位置 table_caption_position', 'option', pos_opts)
        add_row('table_caption_align', '表格标题对齐 table_caption_align', 'option', align_opts)
        add_row('table_caption_template', '表格标题模板 table_caption_template', 'str')

        btns = ctk.CTkFrame(container, fg_color='transparent')
        btns.pack(fill='x', padx=10, pady=(0, 10))

        def reset_defaults():
            d = get_default_export_style()
            for k, v in d.items():
                var = self._vars.get(k)
                if var is None:
                    continue
                try:
                    var.set(v)
                except Exception:
                    try:
                        var.set(str(v))
                    except Exception:
                        pass

        def _style_to_align(v) -> str:
            try:
                if v == WD_ALIGN_PARAGRAPH.CENTER:
                    return 'center'
                if v == WD_ALIGN_PARAGRAPH.RIGHT:
                    return 'right'
                if v == WD_ALIGN_PARAGRAPH.JUSTIFY:
                    return 'justify'
            except Exception:
                pass
            return 'left'

        def _len_pt(v):
            try:
                return float(v.pt)
            except Exception:
                return None

        def _get_east_asia_font(style) -> str:
            try:
                rPr = style.element.rPr
                if rPr is None:
                    return None
                rFonts = rPr.find(qn('w:rFonts'))
                if rFonts is None:
                    return None
                return rFonts.get(qn('w:eastAsia'))
            except Exception:
                return None

        def import_from_docx():
            path = filedialog.askopenfilename(
                title='选择参考 Word 模板',
                filetypes=[('Word 文档', '*.docx')],
            )
            if not path:
                return

            try:
                doc = Document(path)
            except Exception as e:
                messagebox.showerror('导入失败', f'无法读取 Word 文档:\n{e}')
                return

            extracted = {}

            try:
                if doc.sections:
                    sec = doc.sections[0]
                    extracted['margin_top_cm'] = float(sec.top_margin.cm)
                    extracted['margin_bottom_cm'] = float(sec.bottom_margin.cm)
                    extracted['margin_left_cm'] = float(sec.left_margin.cm)
                    extracted['margin_right_cm'] = float(sec.right_margin.cm)
            except Exception:
                pass

            def pick_style(name: str):
                try:
                    return doc.styles[name]
                except Exception:
                    return None

            normal = pick_style('Normal')
            if normal is not None:
                try:
                    if normal.font.name:
                        extracted['body_en'] = normal.font.name
                    if normal.font.size is not None:
                        extracted['body_size_pt'] = float(normal.font.size.pt)
                    ea = _get_east_asia_font(normal)
                    if ea:
                        extracted['body_cn'] = ea
                except Exception:
                    pass

                try:
                    pf = normal.paragraph_format
                    if pf is not None:
                        extracted['body_alignment'] = _style_to_align(pf.alignment)

                        sb = _len_pt(pf.space_before)
                        sa = _len_pt(pf.space_after)
                        fi = _len_pt(pf.first_line_indent)
                        if sb is not None:
                            extracted['body_space_before_pt'] = sb
                        if sa is not None:
                            extracted['body_space_after_pt'] = sa
                        if fi is not None:
                            extracted['body_first_line_indent_pt'] = fi

                        try:
                            if pf.line_spacing_rule == WD_LINE_SPACING.SINGLE:
                                extracted['body_line_spacing'] = 1.0
                            elif pf.line_spacing_rule == WD_LINE_SPACING.ONE_POINT_FIVE:
                                extracted['body_line_spacing'] = 1.5
                            elif pf.line_spacing_rule == WD_LINE_SPACING.DOUBLE:
                                extracted['body_line_spacing'] = 2.0
                            else:
                                ls = pf.line_spacing
                                if isinstance(ls, (int, float)):
                                    extracted['body_line_spacing'] = float(ls)
                        except Exception:
                            pass
                except Exception:
                    pass

            def read_heading(level: int):
                st = pick_style(f'Heading {level}')
                if st is None:
                    return
                try:
                    if st.font.size is not None:
                        extracted[f'heading{level}_size_pt'] = float(st.font.size.pt)
                    extracted[f'heading{level}_bold'] = bool(st.font.bold) if st.font.bold is not None else True
                    ea = _get_east_asia_font(st)
                    if ea:
                        extracted['heading_cn'] = ea
                    if st.font.name:
                        extracted['body_en'] = extracted.get('body_en') or st.font.name
                except Exception:
                    pass
                try:
                    pf = st.paragraph_format
                    if pf is not None:
                        extracted[f'heading{level}_alignment'] = _style_to_align(pf.alignment)
                        sb = _len_pt(pf.space_before)
                        sa = _len_pt(pf.space_after)
                        if sb is not None:
                            extracted[f'heading{level}_space_before_pt'] = sb
                        if sa is not None:
                            extracted[f'heading{level}_space_after_pt'] = sa
                except Exception:
                    pass

            for lv in (1, 2, 3, 4):
                read_heading(lv)

            link = pick_style('Hyperlink')
            if link is not None:
                try:
                    if link.font.size is not None:
                        extracted['hyperlink_size_pt'] = float(link.font.size.pt)
                    if link.font.underline is not None:
                        extracted['hyperlink_underline'] = bool(link.font.underline)
                    if link.font.color is not None and link.font.color.rgb is not None:
                        extracted['hyperlink_color'] = str(link.font.color.rgb)
                except Exception:
                    pass

            caption = pick_style('Caption')
            if caption is not None:
                try:
                    if caption.font.size is not None:
                        extracted['caption_size_pt'] = float(caption.font.size.pt)
                    if caption.font.name:
                        extracted['caption_font'] = caption.font.name
                    ea = _get_east_asia_font(caption)
                    if ea:
                        extracted['body_cn'] = extracted.get('body_cn') or ea
                except Exception:
                    pass

            if not extracted:
                messagebox.showwarning('导入提示', '未从该文档识别到可用的样式信息（可能模板未使用 Word 样式）。')
                return

            for k, v in extracted.items():
                var = self._vars.get(k)
                if var is None:
                    continue
                try:
                    var.set(v)
                except Exception:
                    try:
                        var.set(str(v))
                    except Exception:
                        pass

            try:
                self._style = {**self._style, **extracted}
                self.app.update_status('✅ 已从 Word 模板导入样式（可继续编辑后保存）')
            except Exception:
                pass

        def save_and_close():
            new_style = {}
            for k, var in self._vars.items():
                kind = self._field_types.get(k)
                try:
                    v = var.get()
                except Exception:
                    continue

                if kind == 'bool':
                    new_style[k] = bool(v)
                elif kind == 'float':
                    try:
                        new_style[k] = float(v)
                    except Exception:
                        new_style[k] = self._style.get(k)
                else:
                    new_style[k] = v

            try:
                self.app.config['export_style'] = new_style
                save_config(self.app.config)
                try:
                    self.app.update_status('✅ 已保存导出样式设置')
                except Exception:
                    pass
            except Exception:
                pass

            try:
                self.destroy()
            except Exception:
                pass

        ctk.CTkButton(
            btns,
            text='恢复默认',
            command=reset_defaults,
            fg_color='transparent',
            border_width=1,
            border_color=COLORS['border'],
            text_color=COLORS['text_primary'],
            width=120,
        ).pack(side='left')

        ctk.CTkButton(
            btns,
            text='导入 Word 模板',
            command=import_from_docx,
            fg_color='transparent',
            border_width=1,
            border_color=COLORS['border'],
            text_color=COLORS['text_primary'],
            width=140,
        ).pack(side='left', padx=(8, 0))

        ctk.CTkButton(
            btns,
            text='取消',
            command=self.destroy,
            fg_color='transparent',
            border_width=1,
            border_color=COLORS['border'],
            text_color=COLORS['text_primary'],
            width=90,
        ).pack(side='right', padx=(8, 0))

        ctk.CTkButton(
            btns,
            text='保存',
            command=save_and_close,
            fg_color=COLORS['primary'],
            width=120,
        ).pack(side='right')
