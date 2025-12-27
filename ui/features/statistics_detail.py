# -*- coding: utf-8 -*-
"""字数统计详情功能模块

提供详细的文档统计信息，包括字符数、单词数、段落数、阅读时间等。
"""

import re
import tkinter as tk
from dataclasses import dataclass
from typing import Optional
import customtkinter as ctk

from ui.theme import COLORS


@dataclass
class DocumentStatistics:
    """文档统计数据类"""
    total_chars: int = 0           # 总字符数
    chars_no_spaces: int = 0       # 不含空格字符数
    chinese_chars: int = 0         # 中文字符数
    english_words: int = 0         # 英文单词数
    paragraphs: int = 0            # 段落数
    lines: int = 0                 # 行数
    reading_time_minutes: float = 0.0  # 预计阅读时间（分钟）


class StatisticsDetailFeature:
    """字数统计详情功能类
    
    提供详细的文档统计：
    - 总字符数、不含空格字符数
    - 中文字符数、英文单词数
    - 段落数、行数
    - 预计阅读时间
    """
    
    # 阅读速度常量
    CHINESE_CHARS_PER_MINUTE = 300  # 中文 300 字/分钟
    ENGLISH_WORDS_PER_MINUTE = 200  # 英文 200 词/分钟
    
    def __init__(self, app):
        """初始化统计详情功能
        
        Args:
            app: 主应用实例
        """
        self.app = app
        self._stats: Optional[DocumentStatistics] = None
        self._popup: Optional[ctk.CTkToplevel] = None
    
    def calculate_statistics(self, content: str) -> DocumentStatistics:
        """计算文档统计信息
        
        Args:
            content: 文档内容
            
        Returns:
            DocumentStatistics: 统计数据
        """
        if not content:
            return DocumentStatistics()
        
        # 总字符数
        total_chars = len(content)
        
        # 不含空格字符数
        chars_no_spaces = len(content.replace(' ', '').replace('\t', '').replace('\n', ''))
        
        # 中文字符数
        chinese_chars = self._count_chinese_chars(content)
        
        # 英文单词数
        english_words = self._count_english_words(content)
        
        # 段落数（以空行分隔）
        paragraphs = len([p for p in content.split('\n\n') if p.strip()])
        
        # 行数
        lines = content.count('\n') + 1 if content else 0
        
        # 预计阅读时间
        reading_time = self._calculate_reading_time(chinese_chars, english_words)
        
        self._stats = DocumentStatistics(
            total_chars=total_chars,
            chars_no_spaces=chars_no_spaces,
            chinese_chars=chinese_chars,
            english_words=english_words,
            paragraphs=paragraphs,
            lines=lines,
            reading_time_minutes=reading_time
        )
        
        return self._stats
    
    def _count_chinese_chars(self, text: str) -> int:
        """统计中文字符数
        
        使用 Unicode 范围判断中文字符：
        - CJK 统一汉字: U+4E00 - U+9FFF
        - CJK 扩展 A: U+3400 - U+4DBF
        - CJK 扩展 B-F: U+20000 - U+2FA1F
        
        Args:
            text: 文本内容
            
        Returns:
            int: 中文字符数
        """
        count = 0
        for char in text:
            code = ord(char)
            # CJK 统一汉字
            if 0x4E00 <= code <= 0x9FFF:
                count += 1
            # CJK 扩展 A
            elif 0x3400 <= code <= 0x4DBF:
                count += 1
            # CJK 扩展 B
            elif 0x20000 <= code <= 0x2A6DF:
                count += 1
        return count
    
    def _count_english_words(self, text: str) -> int:
        """统计英文单词数
        
        使用正则表达式匹配英文单词（连续的字母序列）
        
        Args:
            text: 文本内容
            
        Returns:
            int: 英文单词数
        """
        # 匹配英文单词（只包含字母的序列）
        words = re.findall(r'[a-zA-Z]+', text)
        return len(words)
    
    def _calculate_reading_time(self, chinese_chars: int, english_words: int) -> float:
        """计算预计阅读时间
        
        中文 300 字/分钟，英文 200 词/分钟
        
        Args:
            chinese_chars: 中文字符数
            english_words: 英文单词数
            
        Returns:
            float: 阅读时间（分钟）
        """
        chinese_time = chinese_chars / self.CHINESE_CHARS_PER_MINUTE
        english_time = english_words / self.ENGLISH_WORDS_PER_MINUTE
        return round(chinese_time + english_time, 1)
    
    def update_status_bar(self, content: str) -> None:
        """更新状态栏显示
        
        Args:
            content: 文档内容
        """
        stats = self.calculate_statistics(content)
        
        # 格式化阅读时间
        if stats.reading_time_minutes < 1:
            reading_time_str = "< 1 分钟"
        else:
            reading_time_str = f"约 {int(stats.reading_time_minutes)} 分钟"
        
        # 更新状态栏
        try:
            status_text = f"字数: {stats.chars_no_spaces} | 行数: {stats.lines} | 段落: {stats.paragraphs} | 阅读: {reading_time_str}"
            self.app.status_bar_feature.word_count_label.configure(text=status_text)
        except Exception:
            pass
    
    def show_detail_popup(self) -> None:
        """显示详细统计弹窗"""
        # 如果弹窗已存在，聚焦它
        if self._popup and self._popup.winfo_exists():
            self._popup.focus()
            return
        
        # 获取当前内容并计算统计
        try:
            content = self.app.input_text.get("1.0", "end-1c")
        except Exception:
            content = ""
        
        stats = self.calculate_statistics(content)
        
        # 创建弹窗
        self._popup = ctk.CTkToplevel(self.app)
        self._popup.title("📊 文档统计")
        self._popup.geometry("320x380")
        self._popup.resizable(False, False)
        
        # 居中显示
        self._popup.transient(self.app)
        self._popup.grab_set()
        
        # 主容器
        main_frame = ctk.CTkFrame(self._popup, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 标题
        title_label = ctk.CTkLabel(
            main_frame,
            text="📊 文档统计详情",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=COLORS['text_primary']
        )
        title_label.pack(pady=(0, 15))
        
        # 统计项
        stats_items = [
            ("总字符数", str(stats.total_chars)),
            ("不含空格", str(stats.chars_no_spaces)),
            ("中文字符", str(stats.chinese_chars)),
            ("英文单词", str(stats.english_words)),
            ("段落数", str(stats.paragraphs)),
            ("行数", str(stats.lines)),
            ("预计阅读时间", f"{stats.reading_time_minutes} 分钟"),
        ]
        
        for label, value in stats_items:
            row = ctk.CTkFrame(main_frame, fg_color="transparent")
            row.pack(fill="x", pady=4)
            
            label_widget = ctk.CTkLabel(
                row,
                text=label,
                font=ctk.CTkFont(size=13),
                text_color=COLORS['text_secondary'],
                anchor="w"
            )
            label_widget.pack(side="left")
            
            value_widget = ctk.CTkLabel(
                row,
                text=value,
                font=ctk.CTkFont(size=13, weight="bold"),
                text_color=COLORS['text_primary'],
                anchor="e"
            )
            value_widget.pack(side="right")
        
        # 分隔线
        separator = ctk.CTkFrame(main_frame, height=1, fg_color=COLORS['border'])
        separator.pack(fill="x", pady=15)
        
        # 按钮区
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x")
        
        # 复制按钮
        copy_btn = ctk.CTkButton(
            btn_frame,
            text="📋 复制统计",
            width=100,
            height=32,
            corner_radius=8,
            fg_color=COLORS['primary'],
            hover_color=COLORS['primary_hover'],
            command=lambda: self._copy_statistics(stats)
        )
        copy_btn.pack(side="left", padx=(0, 10))
        
        # 关闭按钮
        close_btn = ctk.CTkButton(
            btn_frame,
            text="关闭",
            width=80,
            height=32,
            corner_radius=8,
            fg_color=COLORS['bg_card'],
            text_color=COLORS['text_primary'],
            hover_color=COLORS['highlight'],
            border_width=1,
            border_color=COLORS['border'],
            command=self._popup.destroy
        )
        close_btn.pack(side="right")
    
    def _copy_statistics(self, stats: DocumentStatistics) -> None:
        """复制统计信息到剪贴板
        
        Args:
            stats: 统计数据
        """
        text = f"""文档统计
========
总字符数: {stats.total_chars}
不含空格: {stats.chars_no_spaces}
中文字符: {stats.chinese_chars}
英文单词: {stats.english_words}
段落数: {stats.paragraphs}
行数: {stats.lines}
预计阅读时间: {stats.reading_time_minutes} 分钟"""
        
        try:
            self.app.clipboard_clear()
            self.app.clipboard_append(text)
            self.app.update_status("✅ 统计信息已复制到剪贴板")
        except Exception:
            pass
    
    def bind_status_bar_click(self) -> None:
        """绑定状态栏点击事件"""
        try:
            self.app.status_bar_feature.word_count_label.bind(
                '<Button-1>',
                lambda e: self.show_detail_popup()
            )
            # 添加手型光标
            self.app.status_bar_feature.word_count_label.configure(cursor="hand2")
        except Exception:
            pass
