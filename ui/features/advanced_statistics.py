# -*- coding: utf-8 -*-
"""
增强的统计和分析功能
"""

import re
from collections import Counter
import customtkinter as ctk
from tkinter import messagebox


class AdvancedStatisticsFeature:
    """高级统计分析功能"""
    
    def __init__(self, app):
        self.app = app
        self.stats_dialog = None
        
    def show_advanced_statistics(self):
        """显示高级统计对话框"""
        if self.stats_dialog and self.stats_dialog.winfo_exists():
            self.stats_dialog.focus()
            return
            
        content = self.app.input_text.get("1.0", "end-1c")
        if not content.strip():
            messagebox.showwarning("提示", "文档内容为空！")
            return
            
        # 计算统计数据
        stats = self._calculate_statistics(content)
        
        self.stats_dialog = ctk.CTkToplevel(self.app)
        self.stats_dialog.title("📊 文档统计分析")
        self.stats_dialog.geometry("800x700")
        self.stats_dialog.transient(self.app)
        
        # 标题
        title_label = ctk.CTkLabel(
            self.stats_dialog,
            text="文档统计与分析",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title_label.pack(pady=20)
        
        # 选项卡
        tabview = ctk.CTkTabview(self.stats_dialog, width=760, height=550)
        tabview.pack(padx=20, pady=10)
        
        tabview.add("基础统计")
        tabview.add("可读性分析")
        tabview.add("结构分析")
        tabview.add("关键词")
        
        # === 基础统计 ===
        self._create_basic_stats_tab(tabview.tab("基础统计"), stats)
        
        # === 可读性分析 ===
        self._create_readability_tab(tabview.tab("可读性分析"), stats)
        
        # === 结构分析 ===
        self._create_structure_tab(tabview.tab("结构分析"), stats)
        
        # === 关键词 ===
        self._create_keywords_tab(tabview.tab("关键词"), stats)
        
        # 关闭按钮
        close_btn = ctk.CTkButton(
            self.stats_dialog,
            text="关闭",
            command=self.stats_dialog.destroy,
            width=120
        )
        close_btn.pack(pady=10)
        
    def _calculate_statistics(self, content: str) -> dict:
        """计算各种统计数据"""
        stats = {}
        
        # 基础统计
        stats['total_chars'] = len(content)
        stats['total_chars_no_space'] = len(content.replace(' ', '').replace('\n', ''))
        stats['total_lines'] = content.count('\n') + 1
        stats['total_words'] = len(re.findall(r'\S+', content))
        
        # 中英文统计
        chinese_chars = re.findall(r'[\u4e00-\u9fff]', content)
        english_words = re.findall(r'[a-zA-Z]+', content)
        stats['chinese_chars'] = len(chinese_chars)
        stats['english_words'] = len(english_words)
        
        # Markdown元素统计
        stats['headings'] = self._count_headings(content)
        stats['code_blocks'] = len(re.findall(r'```[\s\S]*?```', content))
        stats['inline_code'] = len(re.findall(r'`[^`]+`', content))
        stats['links'] = len(re.findall(r'\[([^\]]+)\]\([^\)]+\)', content))
        stats['images'] = len(re.findall(r'!\[([^\]]*)\]\([^\)]+\)', content))
        stats['lists'] = len(re.findall(r'^\s*[\*\-\+\d+\.]\s', content, re.MULTILINE))
        stats['tables'] = len(re.findall(r'\|[^\n]+\|', content)) // 2  # 估算表格数
        stats['blockquotes'] = len(re.findall(r'^>', content, re.MULTILINE))
        
        # 段落统计
        paragraphs = [p.strip() for p in content.split('\n\n') if p.strip()]
        stats['paragraphs'] = len(paragraphs)
        
        # 平均值
        if paragraphs:
            avg_para_length = sum(len(p) for p in paragraphs) / len(paragraphs)
            stats['avg_paragraph_length'] = int(avg_para_length)
        else:
            stats['avg_paragraph_length'] = 0
            
        # 句子统计
        sentences = re.split(r'[。！？\.!?]+', content)
        sentences = [s.strip() for s in sentences if s.strip()]
        stats['sentences'] = len(sentences)
        
        if sentences:
            avg_sentence_length = sum(len(s) for s in sentences) / len(sentences)
            stats['avg_sentence_length'] = int(avg_sentence_length)
        else:
            stats['avg_sentence_length'] = 0
            
        # 阅读时间估算（中文250字/分钟，英文200词/分钟）
        reading_time_cn = stats['chinese_chars'] / 250
        reading_time_en = stats['english_words'] / 200
        stats['reading_time'] = max(reading_time_cn, reading_time_en)
        
        # 关键词提取
        stats['keywords'] = self._extract_keywords(content)
        
        return stats
        
    def _count_headings(self, content: str) -> dict:
        """统计各级标题"""
        headings = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0}
        
        for line in content.split('\n'):
            match = re.match(r'^(#{1,6})\s', line)
            if match:
                level = len(match.group(1))
                headings[level] += 1
                
        return headings
        
    def _extract_keywords(self, content: str, top_n: int = 20) -> list:
        """提取关键词"""
        # 移除Markdown标记
        text = re.sub(r'```[\s\S]*?```', '', content)  # 代码块
        text = re.sub(r'`[^`]+`', '', text)  # 行内代码
        text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', text)  # 链接
        text = re.sub(r'[#\*_\-\>\|]', '', text)  # Markdown符号
        
        # 提取词语
        # 中文（简单按字符分）
        chinese_chars = re.findall(r'[\u4e00-\u9fff]+', text)
        chinese_words = []
        for chars in chinese_chars:
            # 简单的2字词提取
            for i in range(len(chars) - 1):
                chinese_words.append(chars[i:i+2])
                
        # 英文单词
        english_words = re.findall(r'[a-zA-Z]{3,}', text.lower())
        
        # 停用词过滤
        stopwords = {'the', 'is', 'at', 'which', 'on', 'a', 'an', 'and', 'or', 'but',
                    '的', '了', '在', '是', '我', '有', '和', '就', '不', '人',
                    '都', '一', '个', '上', '也', '很', '到', '说', '要', '去'}
        
        all_words = chinese_words + english_words
        filtered_words = [w for w in all_words if w not in stopwords and len(w) > 1]
        
        # 统计频率
        word_freq = Counter(filtered_words)
        
        return word_freq.most_common(top_n)
        
    def _create_basic_stats_tab(self, parent, stats):
        """创建基础统计标签页"""
        container = ctk.CTkScrollableFrame(parent, width=720, height=480)
        container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 创建统计卡片
        items = [
            ("📝 总字符数", f"{stats['total_chars']:,}"),
            ("🔤 字符数（不含空格）", f"{stats['total_chars_no_space']:,}"),
            ("📄 总行数", f"{stats['total_lines']:,}"),
            ("💬 总词数", f"{stats['total_words']:,}"),
            ("🈳 中文字符", f"{stats['chinese_chars']:,}"),
            ("🔤 英文单词", f"{stats['english_words']:,}"),
            ("📋 段落数", f"{stats['paragraphs']:,}"),
            ("📝 句子数", f"{stats['sentences']:,}"),
            ("📏 平均段落长度", f"{stats['avg_paragraph_length']} 字符"),
            ("📐 平均句子长度", f"{stats['avg_sentence_length']} 字符"),
            ("⏱️ 预计阅读时间", f"{stats['reading_time']:.1f} 分钟"),
        ]
        
        for i, (label, value) in enumerate(items):
            row = i // 2
            col = i % 2
            
            card = ctk.CTkFrame(container, fg_color="#F9FAFB", corner_radius=10)
            card.grid(row=row, column=col, padx=10, pady=10, sticky="ew")
            
            ctk.CTkLabel(
                card,
                text=label,
                font=ctk.CTkFont(size=12),
                anchor="w"
            ).pack(anchor="w", padx=15, pady=(10, 5))
            
            ctk.CTkLabel(
                card,
                text=value,
                font=ctk.CTkFont(size=18, weight="bold"),
                anchor="w"
            ).pack(anchor="w", padx=15, pady=(0, 10))
            
        # 配置列权重
        container.grid_columnconfigure(0, weight=1)
        container.grid_columnconfigure(1, weight=1)
        
    def _create_readability_tab(self, parent, stats):
        """创建可读性分析标签页"""
        container = ctk.CTkFrame(parent)
        container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 可读性评分（简化版）
        score = self._calculate_readability_score(stats)
        
        # 评分显示
        score_label = ctk.CTkLabel(
            container,
            text=f"可读性评分: {score}/100",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        score_label.pack(pady=20)
        
        # 进度条
        progress = ctk.CTkProgressBar(container, width=400, height=20)
        progress.pack(pady=10)
        progress.set(score / 100)
        
        # 评级
        if score >= 80:
            grade = "优秀 😊"
            color = "#10B981"
        elif score >= 60:
            grade = "良好 🙂"
            color = "#F59E0B"
        elif score >= 40:
            grade = "一般 😐"
            color = "#F97316"
        else:
            grade = "需要改进 😟"
            color = "#EF4444"
            
        grade_label = ctk.CTkLabel(
            container,
            text=grade,
            font=ctk.CTkFont(size=18),
            text_color=color
        )
        grade_label.pack(pady=10)
        
        # 详细分析
        analysis_frame = ctk.CTkFrame(container, fg_color="#F9FAFB")
        analysis_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        analysis_text = ctk.CTkTextbox(analysis_frame, height=300)
        analysis_text.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 生成分析报告
        report = self._generate_readability_report(stats, score)
        analysis_text.insert("1.0", report)
        analysis_text.configure(state="disabled")
        
    def _calculate_readability_score(self, stats) -> int:
        """计算可读性评分（简化算法）"""
        score = 100
        
        # 句子长度惩罚（太长不好读）
        if stats['avg_sentence_length'] > 50:
            score -= (stats['avg_sentence_length'] - 50) * 0.3
        
        # 段落长度惩罚
        if stats['avg_paragraph_length'] > 200:
            score -= (stats['avg_paragraph_length'] - 200) * 0.1
            
        # 结构完整性加分
        if stats['headings']:
            score += 5
        if stats['paragraphs'] > 3:
            score += 5
            
        return max(0, min(100, int(score)))
        
    def _generate_readability_report(self, stats, score) -> str:
        """生成可读性报告"""
        report = "📊 可读性分析报告\n\n"
        
        report += f"总体评分: {score}/100\n\n"
        
        report += "详细分析:\n\n"
        
        # 句子长度分析
        if stats['avg_sentence_length'] > 50:
            report += "❌ 句子偏长，建议将长句拆分为多个短句，提升可读性。\n"
        elif stats['avg_sentence_length'] < 10:
            report += "⚠️ 句子偏短，可以适当合并一些短句。\n"
        else:
            report += "✅ 句子长度适中。\n"
            
        # 段落分析
        if stats['avg_paragraph_length'] > 300:
            report += "❌ 段落过长，建议分段以提高可读性。\n"
        elif stats['paragraphs'] < 3:
            report += "⚠️ 段落较少，内容可能不够丰富。\n"
        else:
            report += "✅ 段落划分合理。\n"
            
        # 结构分析
        total_headings = sum(stats['headings'].values())
        if total_headings == 0:
            report += "⚠️ 缺少标题，建议添加标题以组织内容。\n"
        else:
            report += f"✅ 包含 {total_headings} 个标题，结构清晰。\n"
            
        # 阅读时间
        report += f"\n⏱️ 预计阅读时间: {stats['reading_time']:.1f} 分钟\n"
        
        return report
        
    def _create_structure_tab(self, parent, stats):
        """创建结构分析标签页"""
        container = ctk.CTkScrollableFrame(parent, width=720, height=480)
        container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Markdown元素统计
        ctk.CTkLabel(
            container,
            text="📝 Markdown 元素分布",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", pady=10)
        
        elements = [
            ("📌 代码块", stats['code_blocks']),
            ("💻 行内代码", stats['inline_code']),
            ("🔗 链接", stats['links']),
            ("🖼️ 图片", stats['images']),
            ("📝 列表项", stats['lists']),
            ("📊 表格", stats['tables']),
            ("💬 引用块", stats['blockquotes']),
        ]
        
        for label, count in elements:
            self._create_stat_bar(container, label, count, max(1, max(e[1] for e in elements)))
            
        # 标题层级分布
        ctk.CTkLabel(
            container,
            text="\n📚 标题层级分布",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", pady=10)
        
        max_heading = max(stats['headings'].values()) if stats['headings'].values() else 1
        for level in range(1, 7):
            count = stats['headings'].get(level, 0)
            self._create_stat_bar(container, f"H{level} 标题", count, max(1, max_heading))
            
    def _create_stat_bar(self, parent, label, value, max_value):
        """创建统计条"""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", pady=5)
        
        # 标签
        label_text = ctk.CTkLabel(
            frame,
            text=f"{label}: {value}",
            width=200,
            anchor="w"
        )
        label_text.pack(side="left", padx=5)
        
        # 进度条
        if max_value > 0:
            progress = ctk.CTkProgressBar(frame, width=400, height=15)
            progress.pack(side="left", padx=10)
            progress.set(value / max_value)
            
    def _create_keywords_tab(self, parent, stats):
        """创建关键词标签页"""
        container = ctk.CTkFrame(parent)
        container.pack(fill="both", expand=True, padx=10, pady=10)
        
        ctk.CTkLabel(
            container,
            text="🔑 高频关键词（Top 20）",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=10)
        
        # 关键词列表
        keywords_frame = ctk.CTkScrollableFrame(container, width=700, height=420)
        keywords_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        if stats['keywords']:
            max_freq = stats['keywords'][0][1] if stats['keywords'] else 1
            
            for i, (word, freq) in enumerate(stats['keywords'], 1):
                word_frame = ctk.CTkFrame(keywords_frame, fg_color="#F9FAFB", corner_radius=8)
                word_frame.pack(fill="x", pady=5, padx=5)
                
                # 排名
                rank_label = ctk.CTkLabel(
                    word_frame,
                    text=f"{i}.",
                    width=30,
                    font=ctk.CTkFont(size=14, weight="bold")
                )
                rank_label.pack(side="left", padx=10, pady=10)
                
                # 关键词
                keyword_label = ctk.CTkLabel(
                    word_frame,
                    text=word,
                    width=150,
                    anchor="w",
                    font=ctk.CTkFont(size=13)
                )
                keyword_label.pack(side="left", padx=5, pady=10)
                
                # 频率条
                freq_bar = ctk.CTkProgressBar(word_frame, width=300, height=12)
                freq_bar.pack(side="left", padx=10, pady=10)
                freq_bar.set(freq / max_freq)
                
                # 次数
                freq_label = ctk.CTkLabel(
                    word_frame,
                    text=f"{freq} 次",
                    width=60
                )
                freq_label.pack(side="left", padx=5, pady=10)
        else:
            ctk.CTkLabel(
                keywords_frame,
                text="未找到有效关键词",
                text_color="#9CA3AF"
            ).pack(pady=20)
