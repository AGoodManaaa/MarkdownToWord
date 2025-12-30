# -*- coding: utf-8 -*-
"""
文档统计分析功能
提供字数、段落、标题统计，阅读时间估算，结构分析等
"""

import re
import customtkinter as ctk
from typing import Dict, List, Tuple, Optional
from ui.dialog_utils import set_dialog_icon


class DocumentStatsFeature:
    """文档统计分析功能"""
    
    def __init__(self, app):
        self.app = app
        self.stats_window = None
        
        # 阅读速度（字/分钟）
        self.reading_speed_cn = 300  # 中文
        self.reading_speed_en = 200  # 英文（词/分钟）
    
    def show_stats(self):
        """显示文档统计窗口"""
        if self.stats_window is not None and self.stats_window.winfo_exists():
            self.stats_window.focus()
            return
        
        # 获取当前文档内容
        text = self._get_document_text()
        if not text:
            self._show_message("提示", "文档为空，无法统计。")
            return
        
        # 计算统计数据
        stats = self._calculate_stats(text)
        
        # 创建统计窗口
        self._create_stats_window(stats)
    
    def _get_document_text(self) -> str:
        """获取当前文档文本"""
        try:
            if hasattr(self.app, 'input_text') and self.app.input_text is not None:
                textbox = getattr(self.app.input_text, '_textbox', self.app.input_text)
                return textbox.get("1.0", "end-1c")
        except Exception:
            pass
        return ""
    
    def _calculate_stats(self, text: str) -> Dict:
        """计算文档各项统计数据"""
        lines = text.split('\n')
        
        # 基础统计
        total_chars = len(text)
        total_chars_no_space = len(text.replace(' ', '').replace('\n', '').replace('\t', ''))
        
        # 中英文字符统计
        cn_chars = len(re.findall(r'[\u4e00-\u9fff]', text))
        en_words = len(re.findall(r'\b[a-zA-Z]+\b', text))
        
        # 行和段落统计
        total_lines = len(lines)
        non_empty_lines = len([l for l in lines if l.strip()])
        paragraphs = self._count_paragraphs(text)
        
        # 标题统计
        headings = self._extract_headings(lines)
        heading_counts = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0}
        for level, _ in headings:
            if level in heading_counts:
                heading_counts[level] += 1
        total_headings = sum(heading_counts.values())
        
        # 特殊元素统计
        code_blocks = len(re.findall(r'```[\s\S]*?```', text))
        inline_codes = len(re.findall(r'`[^`\n]+`', text))
        images = len(re.findall(r'!\[.*?\]\(.*?\)', text))
        links = len(re.findall(r'\[.*?\]\(.*?\)', text)) - images
        tables = text.count('|---') + text.count('| ---')
        lists = len(re.findall(r'^[\s]*[-*+]\s', text, re.MULTILINE))
        numbered_lists = len(re.findall(r'^[\s]*\d+\.\s', text, re.MULTILINE))
        
        # 脚注和尾注
        footnotes = len(re.findall(r'\[\^[^\]]+\]:', text))
        endnotes = len(re.findall(r'\[\^\^[^\]]+\]:', text))
        
        # 阅读时间估算
        reading_time_cn = cn_chars / self.reading_speed_cn
        reading_time_en = en_words / self.reading_speed_en
        total_reading_time = reading_time_cn + reading_time_en
        
        return {
            'total_chars': total_chars,
            'total_chars_no_space': total_chars_no_space,
            'cn_chars': cn_chars,
            'en_words': en_words,
            'total_lines': total_lines,
            'non_empty_lines': non_empty_lines,
            'paragraphs': paragraphs,
            'headings': headings,
            'heading_counts': heading_counts,
            'total_headings': total_headings,
            'code_blocks': code_blocks,
            'inline_codes': inline_codes,
            'images': images,
            'links': links,
            'tables': tables,
            'lists': lists + numbered_lists,
            'footnotes': footnotes,
            'endnotes': endnotes,
            'reading_time': total_reading_time,
        }
    
    def _count_paragraphs(self, text: str) -> int:
        """统计段落数量"""
        # 连续的非空行算一个段落
        paragraphs = 0
        in_paragraph = False
        for line in text.split('\n'):
            if line.strip():
                if not in_paragraph:
                    paragraphs += 1
                    in_paragraph = True
            else:
                in_paragraph = False
        return paragraphs
    
    def _extract_headings(self, lines: List[str]) -> List[Tuple[int, str]]:
        """提取标题列表"""
        headings = []
        for line in lines:
            match = re.match(r'^(#{1,6})\s+(.+)', line.strip())
            if match:
                level = len(match.group(1))
                title = match.group(2).strip()
                # 移除可能的锚点
                title = re.sub(r'\s*\{#[^}]+\}\s*$', '', title)
                headings.append((level, title))
        return headings
    
    def _create_stats_window(self, stats: Dict):
        """创建统计窗口"""
        self.stats_window = ctk.CTkToplevel(self.app)
        self.stats_window.title("📊 文档统计分析")
        self.stats_window.geometry("550x650")
        self.stats_window.transient(self.app)
        set_dialog_icon(self.stats_window)
        
        # 居中显示
        self.stats_window.update_idletasks()
        x = self.app.winfo_x() + (self.app.winfo_width() - 550) // 2
        y = self.app.winfo_y() + (self.app.winfo_height() - 650) // 2
        self.stats_window.geometry(f"+{x}+{y}")
        
        # 主容器
        main_frame = ctk.CTkScrollableFrame(self.stats_window)
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 阅读时间（醒目显示）
        time_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray20"))
        time_frame.pack(fill="x", pady=(0, 15))
        
        reading_mins = int(stats['reading_time'])
        reading_secs = int((stats['reading_time'] - reading_mins) * 60)
        time_text = f"{reading_mins} 分 {reading_secs} 秒" if reading_mins > 0 else f"{reading_secs} 秒"
        
        ctk.CTkLabel(
            time_frame, 
            text="⏱️ 预计阅读时间", 
            font=("", 14)
        ).pack(pady=(10, 5))
        ctk.CTkLabel(
            time_frame, 
            text=time_text, 
            font=("", 28, "bold"),
            text_color=("green", "lightgreen")
        ).pack(pady=(0, 10))
        
        # 基础统计
        self._add_section(main_frame, "📝 基础统计", [
            ("总字符数", f"{stats['total_chars']:,}"),
            ("字符数（不含空格）", f"{stats['total_chars_no_space']:,}"),
            ("中文字符", f"{stats['cn_chars']:,}"),
            ("英文单词", f"{stats['en_words']:,}"),
        ])
        
        # 结构统计
        self._add_section(main_frame, "📄 结构统计", [
            ("总行数", f"{stats['total_lines']:,}"),
            ("非空行数", f"{stats['non_empty_lines']:,}"),
            ("段落数", f"{stats['paragraphs']:,}"),
            ("标题数", f"{stats['total_headings']:,}"),
        ])
        
        # 标题层级
        heading_items = []
        for level in range(1, 7):
            count = stats['heading_counts'].get(level, 0)
            if count > 0:
                heading_items.append((f"H{level} 标题", str(count)))
        if heading_items:
            self._add_section(main_frame, "📑 标题层级", heading_items)
        
        # 元素统计
        elements = []
        if stats['code_blocks'] > 0:
            elements.append(("代码块", str(stats['code_blocks'])))
        if stats['inline_codes'] > 0:
            elements.append(("行内代码", str(stats['inline_codes'])))
        if stats['images'] > 0:
            elements.append(("图片", str(stats['images'])))
        if stats['links'] > 0:
            elements.append(("链接", str(stats['links'])))
        if stats['tables'] > 0:
            elements.append(("表格", str(stats['tables'])))
        if stats['lists'] > 0:
            elements.append(("列表项", str(stats['lists'])))
        if stats['footnotes'] > 0:
            elements.append(("脚注", str(stats['footnotes'])))
        if stats['endnotes'] > 0:
            elements.append(("尾注", str(stats['endnotes'])))
        
        if elements:
            self._add_section(main_frame, "🔗 元素统计", elements)
        
        # 文档结构（标题大纲）
        if stats['headings']:
            outline_frame = ctk.CTkFrame(main_frame)
            outline_frame.pack(fill="x", pady=(10, 5))
            
            ctk.CTkLabel(
                outline_frame, 
                text="🗂️ 文档大纲", 
                font=("", 14, "bold"),
                anchor="w"
            ).pack(fill="x", padx=10, pady=5)
            
            outline_text = ctk.CTkTextbox(outline_frame, height=150, wrap="none")
            outline_text.pack(fill="x", padx=10, pady=(0, 10))
            
            for level, title in stats['headings'][:20]:  # 最多显示20个
                indent = "  " * (level - 1)
                outline_text.insert("end", f"{indent}{'#' * level} {title}\n")
            
            if len(stats['headings']) > 20:
                outline_text.insert("end", f"\n... 还有 {len(stats['headings']) - 20} 个标题")
            
            outline_text.configure(state="disabled")
        
        # 关闭按钮
        ctk.CTkButton(
            main_frame, 
            text="关闭", 
            command=self.stats_window.destroy,
            width=120
        ).pack(pady=15)
    
    def _add_section(self, parent, title: str, items: List[Tuple[str, str]]):
        """添加统计区块"""
        frame = ctk.CTkFrame(parent)
        frame.pack(fill="x", pady=(10, 5))
        
        # 标题
        ctk.CTkLabel(
            frame, 
            text=title, 
            font=("", 14, "bold"),
            anchor="w"
        ).pack(fill="x", padx=10, pady=5)
        
        # 统计项（两列布局）
        grid_frame = ctk.CTkFrame(frame, fg_color="transparent")
        grid_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        for i, (label, value) in enumerate(items):
            row = i // 2
            col = i % 2
            
            item_frame = ctk.CTkFrame(grid_frame, fg_color="transparent")
            item_frame.grid(row=row, column=col, sticky="ew", padx=5, pady=2)
            grid_frame.columnconfigure(col, weight=1)
            
            ctk.CTkLabel(
                item_frame, 
                text=f"{label}:", 
                anchor="w",
                text_color=("gray50", "gray70")
            ).pack(side="left")
            ctk.CTkLabel(
                item_frame, 
                text=value, 
                anchor="e",
                font=("", 13, "bold")
            ).pack(side="right")
    
    def _show_message(self, title: str, message: str):
        """显示消息对话框"""
        try:
            from tkinter import messagebox
            messagebox.showinfo(title, message)
        except Exception:
            print(f"{title}: {message}")
    
    def get_quick_stats(self) -> str:
        """获取快速统计信息（用于状态栏）"""
        text = self._get_document_text()
        if not text:
            return "0 字"
        
        cn_chars = len(re.findall(r'[\u4e00-\u9fff]', text))
        en_words = len(re.findall(r'\b[a-zA-Z]+\b', text))
        
        if cn_chars > 0 and en_words > 0:
            return f"{cn_chars} 字 / {en_words} words"
        elif cn_chars > 0:
            return f"{cn_chars} 字"
        else:
            return f"{en_words} words"
