# -*- coding: utf-8 -*-
"""
AI 辅助功能 - 文法检查、摘要生成、翻译等
"""

import threading
from tkinter import messagebox
import customtkinter as ctk
from typing import Optional


class AIAssistantFeature:
    """AI辅助功能管理器"""
    
    def __init__(self, app):
        self.app = app
        self.ai_dialog = None
        self.api_key = None  # TODO: 从配置加载
        self.api_provider = "openai"  # openai, gemini, local等
        
    def show_ai_dialog(self):
        """显示AI助手对话框"""
        if self.ai_dialog and self.ai_dialog.winfo_exists():
            self.ai_dialog.focus()
            return
            
        self.ai_dialog = ctk.CTkToplevel(self.app)
        self.ai_dialog.title("🤖 AI 助手")
        self.ai_dialog.geometry("600x700")
        self.ai_dialog.transient(self.app)
        
        # 标题
        title_label = ctk.CTkLabel(
            self.ai_dialog,
            text="AI 写作助手",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title_label.pack(pady=20)
        
        # 功能选项卡
        tabview = ctk.CTkTabview(self.ai_dialog, width=560, height=550)
        tabview.pack(padx=20, pady=10)
        
        # 添加各个功能标签页
        tabview.add("文法检查")
        tabview.add("生成摘要")
        tabview.add("翻译")
        tabview.add("优化排版")
        tabview.add("续写")
        
        # === 文法检查标签页 ===
        self._create_grammar_check_tab(tabview.tab("文法检查"))
        
        # === 生成摘要标签页 ===
        self._create_summary_tab(tabview.tab("生成摘要"))
        
        # === 翻译标签页 ===
        self._create_translation_tab(tabview.tab("翻译"))
        
        # === 优化排版标签页 ===
        self._create_formatting_tab(tabview.tab("优化排版"))
        
        # === 续写标签页 ===
        self._create_continue_writing_tab(tabview.tab("续写"))
        
        # 底部设置按钮
        settings_btn = ctk.CTkButton(
            self.ai_dialog,
            text="⚙️ API设置",
            command=self._show_api_settings,
            width=120
        )
        settings_btn.pack(pady=10)
        
    def _create_grammar_check_tab(self, parent):
        """创建文法检查标签页"""
        info_label = ctk.CTkLabel(
            parent,
            text="AI 将检查文档的语法、拼写和表达问题",
            font=ctk.CTkFont(size=12)
        )
        info_label.pack(pady=10)
        
        # 选项
        self.check_spelling_var = ctk.BooleanVar(value=True)
        spelling_check = ctk.CTkCheckBox(
            parent,
            text="拼写检查",
            variable=self.check_spelling_var
        )
        spelling_check.pack(anchor="w", padx=20, pady=5)
        
        self.check_grammar_var = ctk.BooleanVar(value=True)
        grammar_check = ctk.CTkCheckBox(
            parent,
            text="语法检查",
            variable=self.check_grammar_var
        )
        grammar_check.pack(anchor="w", padx=20, pady=5)
        
        self.check_style_var = ctk.BooleanVar(value=True)
        style_check = ctk.CTkCheckBox(
            parent,
            text="风格优化建议",
            variable=self.check_style_var
        )
        style_check.pack(anchor="w", padx=20, pady=5)
        
        # 结果显示
        result_frame = ctk.CTkFrame(parent)
        result_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.grammar_result = ctk.CTkTextbox(result_frame, height=250)
        self.grammar_result.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 执行按钮
        check_btn = ctk.CTkButton(
            parent,
            text="🔍 开始检查",
            command=self._run_grammar_check,
            fg_color="#10B981",
            hover_color="#059669"
        )
        check_btn.pack(pady=10)
        
    def _create_summary_tab(self, parent):
        """创建摘要生成标签页"""
        info_label = ctk.CTkLabel(
            parent,
            text="AI 将为你的文档生成简洁的摘要",
            font=ctk.CTkFont(size=12)
        )
        info_label.pack(pady=10)
        
        # 摘要长度选择
        length_frame = ctk.CTkFrame(parent, fg_color="transparent")
        length_frame.pack(pady=10)
        
        ctk.CTkLabel(length_frame, text="摘要长度:").pack(side="left", padx=5)
        
        self.summary_length_var = ctk.StringVar(value="medium")
        length_options = ctk.CTkSegmentedButton(
            length_frame,
            values=["简短", "中等", "详细"],
            variable=self.summary_length_var
        )
        length_options.pack(side="left", padx=5)
        
        # 结果显示
        result_frame = ctk.CTkFrame(parent)
        result_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.summary_result = ctk.CTkTextbox(result_frame, height=250)
        self.summary_result.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 执行按钮
        generate_btn = ctk.CTkButton(
            parent,
            text="✨ 生成摘要",
            command=self._generate_summary,
            fg_color="#8B5CF6",
            hover_color="#7C3AED"
        )
        generate_btn.pack(pady=10)
        
    def _create_translation_tab(self, parent):
        """创建翻译标签页"""
        info_label = ctk.CTkLabel(
            parent,
            text="翻译你的文档内容",
            font=ctk.CTkFont(size=12)
        )
        info_label.pack(pady=10)
        
        # 语言选择
        lang_frame = ctk.CTkFrame(parent, fg_color="transparent")
        lang_frame.pack(pady=10)
        
        ctk.CTkLabel(lang_frame, text="目标语言:").pack(side="left", padx=5)
        
        self.target_lang_var = ctk.StringVar(value="英文")
        lang_menu = ctk.CTkOptionMenu(
            lang_frame,
            values=["英文", "中文", "日文", "韩文", "法文", "德文", "西班牙文"],
            variable=self.target_lang_var
        )
        lang_menu.pack(side="left", padx=5)
        
        # 翻译选项
        self.keep_formatting_var = ctk.BooleanVar(value=True)
        keep_format_check = ctk.CTkCheckBox(
            parent,
            text="保留Markdown格式",
            variable=self.keep_formatting_var
        )
        keep_format_check.pack(anchor="w", padx=20, pady=5)
        
        # 结果显示
        result_frame = ctk.CTkFrame(parent)
        result_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.translation_result = ctk.CTkTextbox(result_frame, height=250)
        self.translation_result.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 按钮组
        btn_frame = ctk.CTkFrame(parent, fg_color="transparent")
        btn_frame.pack(pady=10)
        
        translate_btn = ctk.CTkButton(
            btn_frame,
            text="🌍 翻译",
            command=self._translate_text,
            fg_color="#3B82F6",
            hover_color="#2563EB",
            width=120
        )
        translate_btn.pack(side="left", padx=5)
        
        apply_btn = ctk.CTkButton(
            btn_frame,
            text="✅ 应用到编辑器",
            command=self._apply_translation,
            width=140
        )
        apply_btn.pack(side="left", padx=5)
        
    def _create_formatting_tab(self, parent):
        """创建优化排版标签页"""
        info_label = ctk.CTkLabel(
            parent,
            text="AI 将优化文档的排版和结构",
            font=ctk.CTkFont(size=12)
        )
        info_label.pack(pady=10)
        
        # 优化选项
        options_frame = ctk.CTkFrame(parent, fg_color="transparent")
        options_frame.pack(fill="x", padx=20, pady=10)
        
        self.fix_headings_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            options_frame,
            text="优化标题层级",
            variable=self.fix_headings_var
        ).pack(anchor="w", pady=3)
        
        self.fix_lists_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            options_frame,
            text="规范列表格式",
            variable=self.fix_lists_var
        ).pack(anchor="w", pady=3)
        
        self.add_toc_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            options_frame,
            text="添加目录",
            variable=self.add_toc_var
        ).pack(anchor="w", pady=3)
        
        # 结果预览
        result_frame = ctk.CTkFrame(parent)
        result_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.formatting_result = ctk.CTkTextbox(result_frame, height=250)
        self.formatting_result.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 执行按钮
        btn_frame = ctk.CTkFrame(parent, fg_color="transparent")
        btn_frame.pack(pady=10)
        
        optimize_btn = ctk.CTkButton(
            btn_frame,
            text="✨ 优化排版",
            command=self._optimize_formatting,
            fg_color="#F59E0B",
            hover_color="#D97706",
            width=120
        )
        optimize_btn.pack(side="left", padx=5)
        
        apply_btn = ctk.CTkButton(
            btn_frame,
            text="✅ 应用",
            command=self._apply_formatting,
            width=100
        )
        apply_btn.pack(side="left", padx=5)
        
    def _create_continue_writing_tab(self, parent):
        """创建续写标签页"""
        info_label = ctk.CTkLabel(
            parent,
            text="AI 将根据现有内容智能续写",
            font=ctk.CTkFont(size=12)
        )
        info_label.pack(pady=10)
        
        # 续写长度
        length_frame = ctk.CTkFrame(parent, fg_color="transparent")
        length_frame.pack(pady=10)
        
        ctk.CTkLabel(length_frame, text="续写长度:").pack(side="left", padx=5)
        
        self.continue_length_var = ctk.StringVar(value="100")
        length_entry = ctk.CTkEntry(
            length_frame,
            textvariable=self.continue_length_var,
            width=80
        )
        length_entry.pack(side="left", padx=5)
        
        ctk.CTkLabel(length_frame, text="字").pack(side="left", padx=5)
        
        # 风格选择
        style_frame = ctk.CTkFrame(parent, fg_color="transparent")
        style_frame.pack(pady=10)
        
        ctk.CTkLabel(style_frame, text="写作风格:").pack(side="left", padx=5)
        
        self.writing_style_var = ctk.StringVar(value="延续当前风格")
        style_menu = ctk.CTkOptionMenu(
            style_frame,
            values=["延续当前风格", "正式", "轻松", "学术", "创意"],
            variable=self.writing_style_var
        )
        style_menu.pack(side="left", padx=5)
        
        # 结果显示
        result_frame = ctk.CTkFrame(parent)
        result_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.continue_result = ctk.CTkTextbox(result_frame, height=250)
        self.continue_result.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 按钮
        btn_frame = ctk.CTkFrame(parent, fg_color="transparent")
        btn_frame.pack(pady=10)
        
        continue_btn = ctk.CTkButton(
            btn_frame,
            text="✍️ 续写",
            command=self._continue_writing,
            fg_color="#EC4899",
            hover_color="#DB2777",
            width=120
        )
        continue_btn.pack(side="left", padx=5)
        
        insert_btn = ctk.CTkButton(
            btn_frame,
            text="➕ 插入到编辑器",
            command=self._insert_continued_text,
            width=140
        )
        insert_btn.pack(side="left", padx=5)
        
    def _show_api_settings(self):
        """显示API设置对话框"""
        settings_dialog = ctk.CTkToplevel(self.ai_dialog)
        settings_dialog.title("API 设置")
        settings_dialog.geometry("500x400")
        settings_dialog.transient(self.ai_dialog)
        
        title = ctk.CTkLabel(
            settings_dialog,
            text="AI API 配置",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        title.pack(pady=20)
        
        # API提供商选择
        provider_frame = ctk.CTkFrame(settings_dialog, fg_color="transparent")
        provider_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(provider_frame, text="API提供商:").pack(anchor="w", pady=5)
        
        provider_var = ctk.StringVar(value=self.api_provider)
        provider_menu = ctk.CTkOptionMenu(
            provider_frame,
            values=["OpenAI", "Google Gemini", "本地模型", "自定义"],
            variable=provider_var,
            width=200
        )
        provider_menu.pack(anchor="w", pady=5)
        
        # API Key
        key_frame = ctk.CTkFrame(settings_dialog, fg_color="transparent")
        key_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(key_frame, text="API Key:").pack(anchor="w", pady=5)
        
        key_entry = ctk.CTkEntry(
            key_frame,
            placeholder_text="请输入你的API Key",
            width=400,
            show="*"
        )
        key_entry.pack(anchor="w", pady=5)
        if self.api_key:
            key_entry.insert(0, self.api_key)
        
        # 说明
        info_text = ctk.CTkTextbox(settings_dialog, height=100, width=460)
        info_text.pack(padx=20, pady=10)
        info_text.insert("1.0", 
            "💡 提示:\n"
            "- OpenAI: 需要 OpenAI API Key\n"
            "- Google Gemini: 需要 Google API Key\n"
            "- 本地模型: 需要先配置本地大模型服务\n"
            "- 请妥善保管你的API Key，不要泄露给他人"
        )
        info_text.configure(state="disabled")
        
        # 保存按钮
        save_btn = ctk.CTkButton(
            settings_dialog,
            text="💾 保存",
            command=lambda: self._save_api_settings(
                provider_var.get(),
                key_entry.get(),
                settings_dialog
            ),
            fg_color="#10B981",
            hover_color="#059669"
        )
        save_btn.pack(pady=20)
        
    def _save_api_settings(self, provider: str, api_key: str, dialog):
        """保存API设置"""
        self.api_provider = provider.lower()
        self.api_key = api_key
        
        # TODO: 保存到配置文件
        
        messagebox.showinfo("成功", "API 设置已保存")
        dialog.destroy()
        
    # === AI功能实现方法 ===
    
    def _run_grammar_check(self):
        """执行文法检查"""
        content = self.app.input_text.get("1.0", "end-1c")
        if not content.strip():
            messagebox.showwarning("提示", "编辑器内容为空！")
            return
            
        self.grammar_result.delete("1.0", "end")
        self.grammar_result.insert("1.0", "正在检查...\n")
        
        # TODO: 调用AI API进行检查
        # 这里是示例代码
        def mock_check():
            import time
            time.sleep(2)
            result = "✅ 文法检查完成\n\n"
            result += "未发现明显的语法错误。\n\n"
            result += "建议:\n"
            result += "- 第3段可以更简洁\n"
            result += "- 第5段标点使用规范\n"
            
            self.grammar_result.delete("1.0", "end")
            self.grammar_result.insert("1.0", result)
            
        thread = threading.Thread(target=mock_check, daemon=True)
        thread.start()
        
    def _generate_summary(self):
        """生成摘要"""
        content = self.app.input_text.get("1.0", "end-1c")
        if not content.strip():
            messagebox.showwarning("提示", "编辑器内容为空！")
            return
            
        self.summary_result.delete("1.0", "end")
        self.summary_result.insert("1.0", "正在生成摘要...\n")
        
        # TODO: 调用AI API生成摘要
        
    def _translate_text(self):
        """翻译文本"""
        content = self.app.input_text.get("1.0", "end-1c")
        if not content.strip():
            messagebox.showwarning("提示", "编辑器内容为空！")
            return
            
        target_lang = self.target_lang_var.get()
        self.translation_result.delete("1.0", "end")
        self.translation_result.insert("1.0", f"正在翻译为{target_lang}...\n")
        
        # TODO: 调用AI API翻译
        
    def _apply_translation(self):
        """应用翻译结果到编辑器"""
        translated = self.translation_result.get("1.0", "end-1c")
        if translated.strip():
            self.app.input_text.delete("1.0", "end")
            self.app.input_text.insert("1.0", translated)
            self.app.on_text_change(None)
            messagebox.showinfo("成功", "翻译结果已应用到编辑器")
            
    def _optimize_formatting(self):
        """优化排版"""
        content = self.app.input_text.get("1.0", "end-1c")
        if not content.strip():
            messagebox.showwarning("提示", "编辑器内容为空！")
            return
            
        self.formatting_result.delete("1.0", "end")
        self.formatting_result.insert("1.0", "正在优化排版...\n")
        
        # TODO: 调用AI API优化
        
    def _apply_formatting(self):
        """应用排版优化"""
        formatted = self.formatting_result.get("1.0", "end-1c")
        if formatted.strip():
            self.app.input_text.delete("1.0", "end")
            self.app.input_text.insert("1.0", formatted)
            self.app.on_text_change(None)
            messagebox.showinfo("成功", "排版优化已应用")
            
    def _continue_writing(self):
        """续写文本"""
        content = self.app.input_text.get("1.0", "end-1c")
        if not content.strip():
            messagebox.showwarning("提示", "编辑器内容为空！")
            return
            
        self.continue_result.delete("1.0", "end")
        self.continue_result.insert("1.0", "正在续写...\n")
        
        # TODO: 调用AI API续写
        
    def _insert_continued_text(self):
        """插入续写内容"""
        continued = self.continue_result.get("1.0", "end-1c")
        if continued.strip():
            self.app.input_text.insert("end", "\n\n" + continued)
            self.app.on_text_change(None)
            messagebox.showinfo("成功", "续写内容已插入")
