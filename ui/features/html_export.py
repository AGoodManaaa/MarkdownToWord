# -*- coding: utf-8 -*-
"""
HTML 导出功能
支持将 Markdown 导出为完整的 HTML 文件，内嵌 CSS 样式
"""

import os
import base64
import re
import customtkinter as ctk
from tkinter import filedialog, messagebox
from typing import Optional, Dict
from dataclasses import dataclass
from ui.dialog_utils import set_dialog_icon

try:
    import markdown
    MARKDOWN_AVAILABLE = True
except ImportError:
    MARKDOWN_AVAILABLE = False


@dataclass
class HTMLExportOptions:
    """HTML 导出选项"""
    theme: str = "github"  # github, typora, dark, notion, academic
    embed_images: bool = True  # 是否内嵌图片为 base64
    include_toc: bool = False  # 是否包含目录
    standalone: bool = True  # 是否生成完整 HTML（包含 head）
    title: str = ""  # 文档标题
    custom_css: str = ""  # 自定义 CSS


class HTMLExportFeature:
    """HTML 导出功能"""
    
    def __init__(self, app):
        self.app = app
        self.dialog = None
        self.options = HTMLExportOptions()
        
        # 主题 CSS
        self._themes = {
            'github': self._get_github_css(),
            'typora': self._get_typora_css(),
            'dark': self._get_dark_css(),
            'notion': self._get_notion_css(),
            'academic': self._get_academic_css(),
        }
    
    def show_dialog(self):
        """显示 HTML 导出对话框"""
        if self.dialog is not None and self.dialog.winfo_exists():
            self.dialog.focus()
            return
        
        self.dialog = ctk.CTkToplevel(self.app)
        self.dialog.title("🌐 导出为 HTML")
        self.dialog.geometry("500x450")
        self.dialog.transient(self.app)
        set_dialog_icon(self.dialog)
        
        # 居中显示
        self.dialog.update_idletasks()
        x = self.app.winfo_x() + (self.app.winfo_width() - 500) // 2
        y = self.app.winfo_y() + (self.app.winfo_height() - 450) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        # 主容器
        main_frame = ctk.CTkFrame(self.dialog)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 标题
        ctk.CTkLabel(
            main_frame, 
            text="🌐 导出为 HTML",
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(pady=(0, 15))
        
        # 文档标题
        title_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        title_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(title_frame, text="文档标题:", width=80, anchor="w").pack(side="left")
        self.title_var = ctk.StringVar(value="")
        ctk.CTkEntry(title_frame, textvariable=self.title_var, width=300).pack(side="left", fill="x", expand=True)
        
        # 主题选择
        theme_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        theme_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(theme_frame, text="样式主题:", width=80, anchor="w").pack(side="left")
        
        self.theme_var = ctk.StringVar(value="github")
        themes = [
            ("GitHub", "github"),
            ("Typora", "typora"),
            ("暗色", "dark"),
            ("Notion", "notion"),
            ("学术", "academic"),
        ]
        
        theme_options = ctk.CTkFrame(theme_frame, fg_color="transparent")
        theme_options.pack(side="left", fill="x")
        for name, value in themes:
            ctk.CTkRadioButton(
                theme_options, text=name, variable=self.theme_var, value=value
            ).pack(side="left", padx=5)
        
        # 选项
        options_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        options_frame.pack(fill="x", pady=10)
        
        self.embed_images_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            options_frame, text="内嵌图片 (Base64)", 
            variable=self.embed_images_var
        ).pack(anchor="w", pady=2)
        
        self.include_toc_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            options_frame, text="包含目录 (TOC)", 
            variable=self.include_toc_var
        ).pack(anchor="w", pady=2)
        
        self.standalone_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            options_frame, text="生成完整 HTML 文件", 
            variable=self.standalone_var
        ).pack(anchor="w", pady=2)
        
        # 自定义 CSS
        css_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        css_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(css_frame, text="自定义 CSS (可选):", anchor="w").pack(anchor="w")
        self.custom_css_text = ctk.CTkTextbox(css_frame, height=80)
        self.custom_css_text.pack(fill="x", pady=5)
        
        # 预览信息
        info_label = ctk.CTkLabel(
            main_frame,
            text="💡 提示: 内嵌图片会增加文件大小，但可以离线查看",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        )
        info_label.pack(pady=10)
        
        # 按钮
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(15, 0))
        
        ctk.CTkButton(
            btn_frame, text="📤 导出", width=100,
            fg_color=("green", "darkgreen"),
            command=self._do_export
        ).pack(side="left", padx=(0, 10))
        
        ctk.CTkButton(
            btn_frame, text="👁️ 预览", width=100,
            command=self._preview
        ).pack(side="left", padx=(0, 10))
        
        ctk.CTkButton(
            btn_frame, text="取消", width=80,
            fg_color="gray",
            command=self.dialog.destroy
        ).pack(side="right")
    
    def _do_export(self):
        """执行导出"""
        # 获取选项
        self.options.title = self.title_var.get() or "Markdown Document"
        self.options.theme = self.theme_var.get()
        self.options.embed_images = self.embed_images_var.get()
        self.options.include_toc = self.include_toc_var.get()
        self.options.standalone = self.standalone_var.get()
        self.options.custom_css = self.custom_css_text.get("1.0", "end-1c")
        
        # 选择保存路径
        file_path = filedialog.asksaveasfilename(
            title="保存 HTML 文件",
            defaultextension=".html",
            filetypes=[("HTML 文件", "*.html"), ("所有文件", "*.*")],
            parent=self.dialog
        )
        
        if not file_path:
            return
        
        try:
            # 获取 Markdown 内容
            content = self.app.input_text.get("1.0", "end-1c")
            
            # 转换为 HTML
            html = self.convert_to_html(content, self.options)
            
            # 保存文件
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(html)
            
            messagebox.showinfo("成功", f"HTML 文件已保存到:\n{file_path}", parent=self.dialog)
            self.dialog.destroy()
            
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {e}", parent=self.dialog)
    
    def _preview(self):
        """预览 HTML"""
        import tempfile
        import webbrowser
        
        try:
            # 获取选项
            self.options.title = self.title_var.get() or "Markdown Document"
            self.options.theme = self.theme_var.get()
            self.options.embed_images = self.embed_images_var.get()
            self.options.include_toc = self.include_toc_var.get()
            self.options.standalone = True
            self.options.custom_css = self.custom_css_text.get("1.0", "end-1c")
            
            # 获取 Markdown 内容
            content = self.app.input_text.get("1.0", "end-1c")
            
            # 转换为 HTML
            html = self.convert_to_html(content, self.options)
            
            # 保存到临时文件
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                f.write(html)
                temp_path = f.name
            
            # 在浏览器中打开
            webbrowser.open(f'file://{temp_path}')
            
        except Exception as e:
            messagebox.showerror("错误", f"预览失败: {e}", parent=self.dialog)
    
    def convert_to_html(self, markdown_content: str, options: HTMLExportOptions = None) -> str:
        """将 Markdown 转换为 HTML"""
        if options is None:
            options = HTMLExportOptions()
        
        # 转换 Markdown 到 HTML
        if MARKDOWN_AVAILABLE:
            extensions = ['tables', 'fenced_code', 'codehilite', 'toc', 'nl2br']
            html_body = markdown.markdown(markdown_content, extensions=extensions)
        else:
            # 简单的 Markdown 转换
            html_body = self._simple_markdown_to_html(markdown_content)
        
        # 处理图片
        if options.embed_images:
            html_body = self._embed_images(html_body)
        
        # 生成目录
        toc_html = ""
        if options.include_toc:
            toc_html = self._generate_toc(markdown_content)
        
        # 获取 CSS
        css = self._themes.get(options.theme, self._themes['github'])
        if options.custom_css:
            css += f"\n/* Custom CSS */\n{options.custom_css}"
        
        if options.standalone:
            return self._wrap_html(html_body, options.title, css, toc_html)
        else:
            return html_body
    
    def _simple_markdown_to_html(self, content: str) -> str:
        """简单的 Markdown 转 HTML（不依赖 markdown 库）"""
        lines = content.split('\n')
        html_lines = []
        in_code_block = False
        in_list = False
        
        for line in lines:
            # 代码块
            if line.strip().startswith('```'):
                if in_code_block:
                    html_lines.append('</code></pre>')
                    in_code_block = False
                else:
                    lang = line.strip()[3:]
                    html_lines.append(f'<pre><code class="language-{lang}">')
                    in_code_block = True
                continue
            
            if in_code_block:
                html_lines.append(self._escape_html(line))
                continue
            
            # 标题
            if line.startswith('#'):
                level = len(line) - len(line.lstrip('#'))
                text = line.lstrip('#').strip()
                html_lines.append(f'<h{level}>{text}</h{level}>')
                continue
            
            # 列表
            if line.strip().startswith(('- ', '* ', '+ ')):
                if not in_list:
                    html_lines.append('<ul>')
                    in_list = True
                text = line.strip()[2:]
                html_lines.append(f'<li>{self._inline_format(text)}</li>')
                continue
            elif in_list and not line.strip():
                html_lines.append('</ul>')
                in_list = False
            
            # 引用
            if line.startswith('>'):
                text = line[1:].strip()
                html_lines.append(f'<blockquote>{self._inline_format(text)}</blockquote>')
                continue
            
            # 分隔线
            if line.strip() in ('---', '***', '___'):
                html_lines.append('<hr>')
                continue
            
            # 普通段落
            if line.strip():
                html_lines.append(f'<p>{self._inline_format(line)}</p>')
            else:
                html_lines.append('')
        
        if in_list:
            html_lines.append('</ul>')
        
        return '\n'.join(html_lines)
    
    def _inline_format(self, text: str) -> str:
        """处理行内格式"""
        # 粗体
        text = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', text)
        text = re.sub(r'__(.+?)__', r'<strong>\1</strong>', text)
        # 斜体
        text = re.sub(r'\*(.+?)\*', r'<em>\1</em>', text)
        text = re.sub(r'_(.+?)_', r'<em>\1</em>', text)
        # 行内代码
        text = re.sub(r'`(.+?)`', r'<code>\1</code>', text)
        # 链接
        text = re.sub(r'\[(.+?)\]\((.+?)\)', r'<a href="\2">\1</a>', text)
        # 图片
        text = re.sub(r'!\[(.+?)\]\((.+?)\)', r'<img src="\2" alt="\1">', text)
        return text
    
    def _escape_html(self, text: str) -> str:
        """转义 HTML 特殊字符"""
        return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    
    def _embed_images(self, html: str) -> str:
        """将图片转换为 base64 内嵌"""
        def replace_image(match):
            src = match.group(1)
            if src.startswith(('http://', 'https://', 'data:')):
                return match.group(0)  # 保持不变
            
            try:
                # 尝试读取本地图片
                if os.path.exists(src):
                    with open(src, 'rb') as f:
                        data = base64.b64encode(f.read()).decode()
                    ext = os.path.splitext(src)[1].lower()
                    mime = {
                        '.png': 'image/png',
                        '.jpg': 'image/jpeg',
                        '.jpeg': 'image/jpeg',
                        '.gif': 'image/gif',
                        '.svg': 'image/svg+xml',
                        '.webp': 'image/webp',
                    }.get(ext, 'image/png')
                    return f'src="data:{mime};base64,{data}"'
            except Exception:
                pass
            
            return match.group(0)
        
        return re.sub(r'src="([^"]+)"', replace_image, html)
    
    def _generate_toc(self, content: str) -> str:
        """生成目录"""
        toc_items = []
        for line in content.split('\n'):
            if line.startswith('#'):
                level = len(line) - len(line.lstrip('#'))
                text = line.lstrip('#').strip()
                anchor = re.sub(r'[^\w\s-]', '', text.lower()).replace(' ', '-')
                indent = '  ' * (level - 1)
                toc_items.append(f'{indent}<li><a href="#{anchor}">{text}</a></li>')
        
        if toc_items:
            return f'<nav class="toc"><h2>目录</h2><ul>{"".join(toc_items)}</ul></nav>'
        return ""
    
    def _wrap_html(self, body: str, title: str, css: str, toc: str = "") -> str:
        """包装为完整 HTML"""
        return f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <style>
{css}
    </style>
</head>
<body>
    <article class="markdown-body">
        {toc}
        {body}
    </article>
</body>
</html>'''
    
    def _get_github_css(self) -> str:
        """GitHub 风格 CSS"""
        return '''
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
            font-size: 16px;
            line-height: 1.6;
            color: #24292f;
            background-color: #ffffff;
            margin: 0;
            padding: 20px;
        }
        .markdown-body {
            max-width: 900px;
            margin: 0 auto;
            padding: 20px 40px;
        }
        h1, h2, h3, h4, h5, h6 {
            margin-top: 24px;
            margin-bottom: 16px;
            font-weight: 600;
            line-height: 1.25;
        }
        h1 { font-size: 2em; border-bottom: 1px solid #d0d7de; padding-bottom: 0.3em; }
        h2 { font-size: 1.5em; border-bottom: 1px solid #d0d7de; padding-bottom: 0.3em; }
        h3 { font-size: 1.25em; }
        a { color: #0969da; text-decoration: none; }
        a:hover { text-decoration: underline; }
        code {
            background-color: rgba(175,184,193,0.2);
            padding: 0.2em 0.4em;
            border-radius: 6px;
            font-family: ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace;
            font-size: 85%;
        }
        pre {
            background-color: #f6f8fa;
            padding: 16px;
            border-radius: 6px;
            overflow-x: auto;
        }
        pre code {
            background-color: transparent;
            padding: 0;
        }
        blockquote {
            margin: 0;
            padding: 0 1em;
            color: #656d76;
            border-left: 0.25em solid #d0d7de;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin: 16px 0;
        }
        th, td {
            border: 1px solid #d0d7de;
            padding: 6px 13px;
        }
        th {
            background-color: #f6f8fa;
            font-weight: 600;
        }
        img { max-width: 100%; }
        hr {
            height: 0.25em;
            padding: 0;
            margin: 24px 0;
            background-color: #d0d7de;
            border: 0;
        }
        .toc {
            background-color: #f6f8fa;
            padding: 16px;
            border-radius: 6px;
            margin-bottom: 24px;
        }
        .toc h2 { margin-top: 0; font-size: 1.2em; }
        .toc ul { padding-left: 20px; }
        .toc a { color: #0969da; }
        '''
    
    def _get_typora_css(self) -> str:
        """Typora 风格 CSS"""
        return '''
        body {
            font-family: "Open Sans", "Clear Sans", "Helvetica Neue", Helvetica, Arial, sans-serif;
            font-size: 16px;
            line-height: 1.8;
            color: #333;
            background-color: #fff;
            margin: 0;
            padding: 20px;
        }
        .markdown-body {
            max-width: 860px;
            margin: 0 auto;
            padding: 30px;
        }
        h1, h2, h3, h4, h5, h6 {
            margin-top: 1rem;
            margin-bottom: 1rem;
            font-weight: bold;
            line-height: 1.4;
        }
        h1 { font-size: 2.25rem; }
        h2 { font-size: 1.75rem; }
        h3 { font-size: 1.5rem; }
        a { color: #4183c4; }
        code {
            background-color: #f3f4f4;
            padding: 2px 4px;
            border-radius: 4px;
            color: #c7254e;
            font-family: Consolas, Monaco, monospace;
        }
        pre {
            background-color: #f8f8f8;
            padding: 1em;
            border-radius: 3px;
            overflow-x: auto;
        }
        blockquote {
            border-left: 4px solid #dfe2e5;
            padding: 0 15px;
            color: #777;
        }
        table {
            border-collapse: collapse;
            width: 100%;
        }
        th, td {
            border: 1px solid #dfe2e5;
            padding: 8px 12px;
        }
        '''
    
    def _get_dark_css(self) -> str:
        """暗色主题 CSS"""
        return '''
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
            font-size: 16px;
            line-height: 1.6;
            color: #d4d4d4;
            background-color: #1e1e1e;
            margin: 0;
            padding: 20px;
        }
        .markdown-body {
            max-width: 900px;
            margin: 0 auto;
            padding: 20px 40px;
        }
        h1, h2, h3, h4, h5, h6 {
            color: #fff;
            margin-top: 24px;
            margin-bottom: 16px;
        }
        h1, h2 { border-bottom: 1px solid #404040; padding-bottom: 0.3em; }
        a { color: #569cd6; }
        code {
            background-color: #2d2d2d;
            color: #ce9178;
            padding: 0.2em 0.4em;
            border-radius: 4px;
            font-family: Consolas, Monaco, monospace;
        }
        pre {
            background-color: #1e1e1e;
            border: 1px solid #404040;
            padding: 16px;
            border-radius: 6px;
        }
        blockquote {
            border-left: 4px solid #569cd6;
            background-color: #2d2d2d;
            padding: 10px 20px;
            color: #9cdcfe;
        }
        table {
            border-collapse: collapse;
            width: 100%;
        }
        th, td {
            border: 1px solid #404040;
            padding: 8px 12px;
        }
        th { background-color: #2d2d2d; }
        hr { background-color: #404040; }
        '''
    
    def _get_notion_css(self) -> str:
        """Notion 风格 CSS"""
        return '''
        body {
            font-family: ui-sans-serif, -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
            font-size: 16px;
            line-height: 1.5;
            color: #37352f;
            background-color: #fff;
            margin: 0;
            padding: 20px;
        }
        .markdown-body {
            max-width: 900px;
            margin: 0 auto;
            padding: 20px 96px;
        }
        h1 { font-size: 1.875rem; font-weight: 700; }
        h2 { font-size: 1.5rem; font-weight: 600; }
        h3 { font-size: 1.25rem; font-weight: 600; }
        a { color: #37352f; text-decoration: underline; }
        code {
            background-color: rgba(135,131,120,0.15);
            color: #eb5757;
            padding: 0.2em 0.4em;
            border-radius: 3px;
            font-family: "SFMono-Regular", Consolas, monospace;
        }
        pre {
            background-color: #f7f6f3;
            padding: 16px;
            border-radius: 3px;
        }
        blockquote {
            border-left: 3px solid #000;
            padding-left: 14px;
            margin-left: 0;
        }
        '''
    
    def _get_academic_css(self) -> str:
        """学术风格 CSS"""
        return '''
        body {
            font-family: Georgia, "Times New Roman", serif;
            font-size: 16px;
            line-height: 1.8;
            color: #111;
            background-color: #fffff8;
            margin: 0;
            padding: 20px;
        }
        .markdown-body {
            max-width: 700px;
            margin: 0 auto;
            padding: 40px;
        }
        h1, h2, h3, h4, h5, h6 {
            font-weight: normal;
            margin-top: 2em;
            margin-bottom: 1em;
        }
        h1 { font-size: 1.5rem; text-align: center; }
        h2 { font-size: 1.3rem; }
        a { color: #0645ad; }
        code {
            background-color: #f5f5f5;
            padding: 2px 4px;
            font-family: "Courier New", monospace;
        }
        pre {
            background-color: #f5f5f5;
            padding: 1em;
            border: 1px solid #ddd;
        }
        blockquote {
            border-left: 2px solid #ccc;
            padding-left: 1em;
            color: #666;
            font-style: italic;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin: 1em 0;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }
        '''
    
    def export_quick(self, file_path: str = None) -> Optional[str]:
        """快速导出（使用默认选项）"""
        if file_path is None:
            file_path = filedialog.asksaveasfilename(
                title="保存 HTML 文件",
                defaultextension=".html",
                filetypes=[("HTML 文件", "*.html"), ("所有文件", "*.*")]
            )
        
        if not file_path:
            return None
        
        try:
            content = self.app.input_text.get("1.0", "end-1c")
            html = self.convert_to_html(content)
            
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(html)
            
            return file_path
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {e}")
            return None
