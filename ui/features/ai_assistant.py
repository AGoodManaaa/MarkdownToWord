# -*- coding: utf-8 -*-
"""
AI 写作助手功能
支持 OpenAI GPT 集成，提供润色、续写、翻译、总结等功能
"""

import os
import json
import threading
import customtkinter as ctk
from tkinter import messagebox, END
from typing import Optional
from dataclasses import dataclass

# 尝试导入 OpenAI
try:
    import openai
    HAS_OPENAI = True
except ImportError:
    HAS_OPENAI = False


@dataclass
class AIConfig:
    """AI 配置"""
    provider: str = "openai"  # openai, deepseek, siliconflow
    api_key: str = ""
    api_base: str = "https://api.openai.com/v1"
    model: str = "gpt-3.5-turbo"
    temperature: float = 0.7
    max_tokens: int = 2000

# 服务商预设配置
PROVIDERS = {
    "openai": {
        "name": "OpenAI",
        "api_base": "https://api.openai.com/v1",
        "models": ["gpt-3.5-turbo", "gpt-4", "gpt-4-turbo", "gpt-4o", "gpt-4o-mini"]
    },
    "deepseek": {
        "name": "DeepSeek (深度求索)",
        "api_base": "https://api.deepseek.com",
        "models": ["deepseek-chat", "deepseek-coder"]
    },
    "siliconflow": {
        "name": "SiliconFlow (硅基流动)",
        "api_base": "https://api.siliconflow.cn/v1",
        "models": [
            "deepseek-ai/DeepSeek-V2-Chat",
            "deepseek-ai/DeepSeek-V2.5",
            "Qwen/Qwen2.5-72B-Instruct",
            "Qwen/Qwen2.5-7B-Instruct",
            "meta-llama/Meta-Llama-3.1-70B-Instruct"
        ]
    }
}


class AIAssistantFeature:
    """AI 写作助手功能"""
    
    # 预设提示词
    PROMPTS = {
        "polish": ("✨ 润色", "请润色以下文本，使其更加流畅、专业，保持原意：\n\n{text}"),
        "continue": ("📝 续写", "请根据以下内容继续写作，保持风格一致，续写约200字：\n\n{text}"),
        "translate_en": ("🌐 译英", "请将以下中文翻译成英文：\n\n{text}"),
        "translate_cn": ("🌐 译中", "请将以下英文翻译成中文：\n\n{text}"),
        "summarize": ("📋 总结", "请用简洁的语言总结以下内容的要点：\n\n{text}"),
        "expand": ("📖 扩写", "请将以下内容扩展详细，增加更多细节：\n\n{text}"),
        "simplify": ("🔤 简化", "请用更简单易懂的语言重写：\n\n{text}"),
        "formal": ("📜 正式化", "请将以下文本改写为正式书面语体：\n\n{text}"),
        "fix_grammar": ("🔧 修正", "请检查并修正语法和拼写错误：\n\n{text}"),
        "casual": ("💬 口语化", "请将以下文本改写为轻松口语风格：\n\n{text}"),
    }
    
    def __init__(self, app):
        self.app = app
        self.dialog = None
        self.config_file = os.path.join(os.path.dirname(__file__), '..', '..', 'ai_config.json')
        self.config = self._load_config()
        self.is_generating = False
    
    def _load_config(self) -> AIConfig:
        """加载配置"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return AIConfig(**data)
        except Exception:
            pass
        return AIConfig()
    
    def _save_config(self):
        """保存配置"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                data = {
                    "provider": getattr(self.config, "provider", "openai"),
                    "api_key": self.config.api_key,
                    "api_base": self.config.api_base,
                    "model": self.config.model,
                    "temperature": self.config.temperature,
                    "max_tokens": self.config.max_tokens
                }
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存AI配置失败: {e}")
    
    def show_dialog(self):
        """显示AI助手对话框"""
        if self.dialog is not None and self.dialog.winfo_exists():
            self.dialog.focus()
            return
        
        self.dialog = ctk.CTkToplevel(self.app)
        self.dialog.title("🤖 AI 写作助手")
        self.dialog.geometry("750x680")
        self.dialog.transient(self.app)
        
        # 居中
        self.dialog.update_idletasks()
        x = self.app.winfo_x() + (self.app.winfo_width() - 750) // 2
        y = self.app.winfo_y() + (self.app.winfo_height() - 680) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        try:
            from ui.dialog_utils import set_dialog_icon
            set_dialog_icon(self.dialog)
        except:
            pass
        
        main_frame = ctk.CTkFrame(self.dialog)
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # API 配置区
        config_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray20"))
        config_frame.pack(fill="x", pady=(0, 10))
        
        # 顶部：服务商选择 + 简易配置
        provider_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        provider_frame.pack(side="left", padx=5)
        
        ctk.CTkLabel(provider_frame, text="服务商:").pack(side="left")
        self.provider_var = ctk.StringVar(value=getattr(self.config, "provider", "openai"))
        
        provider_names = [p["name"] for p in PROVIDERS.values()]
        # 映射显示名到 key
        self.provider_map = {v["name"]: k for k, v in PROVIDERS.items()}
        current_provider_name = PROVIDERS.get(self.provider_var.get(), PROVIDERS["openai"])["name"]
        
        self.provider_menu = ctk.CTkOptionMenu(
            provider_frame, 
            values=provider_names, 
            width=140,
            command=self._on_provider_change
        )
        self.provider_menu.pack(side="left", padx=5)
        self.provider_menu.set(current_provider_name)

        ctk.CTkLabel(config_frame, text="API Key:").pack(side="left", padx=5)
        self.api_key_entry = ctk.CTkEntry(config_frame, width=200, show="*")
        self.api_key_entry.pack(side="left", padx=5)
        if self.config.api_key:
            self.api_key_entry.insert(0, self.config.api_key)
        
        ctk.CTkLabel(config_frame, text="模型:").pack(side="left", padx=5)
        self.model_var = ctk.StringVar(value=self.config.model)
        
        # 获取当前服务商的模型列表
        current_models = PROVIDERS.get(self.provider_var.get(), PROVIDERS["openai"])["models"]
        
        self.model_menu = ctk.CTkComboBox(config_frame, values=current_models, variable=self.model_var, width=160)
        self.model_menu.pack(side="left")
        
        ctk.CTkButton(config_frame, text="⚙️", width=35, command=self._show_settings).pack(side="right", padx=10)
        
        # 功能按钮
        func_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        func_frame.pack(fill="x", pady=(0, 10))
        
        ctk.CTkLabel(func_frame, text="功能:", font=("", 13, "bold")).pack(side="left", padx=(0, 8))
        
        btn_container = ctk.CTkFrame(func_frame, fg_color="transparent")
        btn_container.pack(side="left", fill="x", expand=True)
        
        row1 = ctk.CTkFrame(btn_container, fg_color="transparent")
        row1.pack(fill="x")
        row2 = ctk.CTkFrame(btn_container, fg_color="transparent")
        row2.pack(fill="x", pady=(4, 0))
        
        prompts = list(self.PROMPTS.items())
        for i, (key, (name, _)) in enumerate(prompts[:5]):
            ctk.CTkButton(row1, text=name, width=85, command=lambda k=key: self._execute(k)).pack(side="left", padx=2)
        for i, (key, (name, _)) in enumerate(prompts[5:]):
            ctk.CTkButton(row2, text=name, width=85, command=lambda k=key: self._execute(k)).pack(side="left", padx=2)
        
        # 输入区
        ctk.CTkLabel(main_frame, text="输入文本:", anchor="w").pack(fill="x")
        self.input_text = ctk.CTkTextbox(main_frame, height=140)
        self.input_text.pack(fill="x", pady=(5, 10))
        self._load_selection()
        
        # 自定义提示词
        custom_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        custom_frame.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(custom_frame, text="自定义指令:").pack(side="left")
        self.custom_prompt = ctk.CTkEntry(custom_frame, placeholder_text="例如：改写成诗歌形式...")
        self.custom_prompt.pack(side="left", fill="x", expand=True, padx=5)
        ctk.CTkButton(custom_frame, text="▶ 执行", width=70, command=self._execute_custom).pack(side="right")
        
        # 输出区
        output_header = ctk.CTkFrame(main_frame, fg_color="transparent")
        output_header.pack(fill="x")
        ctk.CTkLabel(output_header, text="AI 输出:", anchor="w").pack(side="left")
        self.status_label = ctk.CTkLabel(output_header, text="", text_color=("gray50", "gray70"))
        self.status_label.pack(side="right")
        
        self.output_text = ctk.CTkTextbox(main_frame, height=180)
        self.output_text.pack(fill="both", expand=True, pady=(5, 10))
        
        # 底部按钮
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x")
        
        ctk.CTkButton(btn_frame, text="📋 复制", width=80, command=self._copy).pack(side="left", padx=(0, 5))
        ctk.CTkButton(btn_frame, text="📥 插入", width=80, fg_color=("green", "darkgreen"), command=self._insert).pack(side="left", padx=(0, 5))
        ctk.CTkButton(btn_frame, text="🔄 替换选中", width=100, command=self._replace).pack(side="left")
        
        self.stop_btn = ctk.CTkButton(btn_frame, text="⏹ 停止", width=70, fg_color=("red", "darkred"), command=self._stop, state="disabled")
        self.stop_btn.pack(side="right")
    
    def _load_selection(self):
        """加载选中文本"""
        try:
            textbox = self._get_textbox()
            if textbox:
                try:
                    selected = textbox.get("sel.first", "sel.last")
                    if selected:
                        self.input_text.insert("1.0", selected)
                except:
                    pass
        except:
            pass
    
    def _get_textbox(self):
        try:
            if hasattr(self.app, 'input_text') and self.app.input_text is not None:
                return getattr(self.app.input_text, '_textbox', self.app.input_text)
        except:
            pass
        return None
    
    def _show_settings(self):
        """高级设置"""
        settings = ctk.CTkToplevel(self.dialog)
        settings.title("⚙️ 高级设置")
        settings.geometry("420x320")
        settings.transient(self.dialog)
        settings.grab_set()
        
        frame = ctk.CTkFrame(settings)
        frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        ctk.CTkLabel(frame, text="API Base URL:").pack(anchor="w")
        base_entry = ctk.CTkEntry(frame, width=380)
        base_entry.pack(fill="x", pady=(0, 10))
        base_entry.insert(0, self.config.api_base)
        
        ctk.CTkLabel(frame, text=f"Temperature: {self.config.temperature}").pack(anchor="w")
        temp_slider = ctk.CTkSlider(frame, from_=0, to=1, number_of_steps=10)
        temp_slider.pack(fill="x", pady=(0, 10))
        temp_slider.set(self.config.temperature)
        
        ctk.CTkLabel(frame, text="Max Tokens:").pack(anchor="w")
        tokens_entry = ctk.CTkEntry(frame, width=100)
        tokens_entry.pack(anchor="w", pady=(0, 10))
        tokens_entry.insert(0, str(self.config.max_tokens))
        
        def save():
            self.config.api_base = base_entry.get()
            self.config.temperature = temp_slider.get()
            try:
                self.config.max_tokens = int(tokens_entry.get())
            except:
                pass
            self._save_config()
            settings.destroy()
        
        ctk.CTkButton(frame, text="保存", command=save).pack(pady=15)
    
    def _on_provider_change(self, provider_name: str):
        """服务商切换处理"""
        provider_key = self.provider_map.get(provider_name)
        if not provider_key:
            return
            
        # 更新配置
        self.config.provider = provider_key
        provider_data = PROVIDERS[provider_key]
        self.config.api_base = provider_data["api_base"]
        
        # 更新模型列表
        models = provider_data["models"]
        self.model_menu.configure(values=models)
        self.model_var.set(models[0])
        self.config.model = models[0]
        
        # 提示更新 API Key
        self.api_key_entry.delete(0, 9999) # 使用 9999 而不是 END 因为 END 在这里可能未导入
        # messagebox.showinfo("提示", f"已切换到 {provider_name}。\n请确保输入正确的 API Key。")

    def _execute(self, prompt_key: str):
        """执行预设提示"""
        text = self.input_text.get("1.0", "end-1c").strip()
        if not text:
            messagebox.showwarning("警告", "请输入要处理的文本")
            return
        
        _, prompt_template = self.PROMPTS[prompt_key]
        prompt = prompt_template.format(text=text)
        self._call_api(prompt)
    
    def _execute_custom(self):
        """执行自定义提示"""
        text = self.input_text.get("1.0", "end-1c").strip()
        custom = self.custom_prompt.get().strip()
        if not text:
            messagebox.showwarning("警告", "请输入要处理的文本")
            return
        if not custom:
            messagebox.showwarning("警告", "请输入自定义指令")
            return
        
        prompt = f"{custom}\n\n{text}"
        self._call_api(prompt)
    
    def _call_api(self, prompt: str):
        """调用 API"""
        if not HAS_OPENAI:
            messagebox.showerror("错误", "请先安装 openai 库:\npip install openai")
            return
        
        self.config.api_key = self.api_key_entry.get()
        self.config.model = self.model_var.get()
        self._save_config()
        
        if not self.config.api_key:
            messagebox.showwarning("警告", "请输入 API Key")
            return
        
        self.output_text.delete("1.0", END)
        self.status_label.configure(text="⏳ 正在生成...")
        self.stop_btn.configure(state="normal")
        self.is_generating = True
        
        thread = threading.Thread(target=self._api_thread, args=(prompt,), daemon=True)
        thread.start()
    
    def _api_thread(self, prompt: str):
        """API 调用线程"""
        try:
            client = openai.OpenAI(api_key=self.config.api_key, base_url=self.config.api_base)
            
            response = client.chat.completions.create(
                model=self.config.model,
                messages=[
                    {"role": "system", "content": "你是一个专业的写作助手。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=self.config.temperature,
                max_tokens=self.config.max_tokens,
                stream=True
            )
            
            for chunk in response:
                if not self.is_generating:
                    break
                if chunk.choices[0].delta.content:
                    content = chunk.choices[0].delta.content
                    self._append(content)
            
            self._done()
            
        except Exception as e:
            self._error(str(e))
    
    def _append(self, text: str):
        try:
            self.dialog.after(0, lambda: self.output_text.insert(END, text))
        except:
            pass
    
    def _done(self):
        try:
            self.dialog.after(0, lambda: self.status_label.configure(text="✅ 完成"))
            self.dialog.after(0, lambda: self.stop_btn.configure(state="disabled"))
        except:
            pass
        self.is_generating = False
    
    def _error(self, msg: str):
        try:
            self.dialog.after(0, lambda: self.status_label.configure(text=f"❌ 错误"))
            self.dialog.after(0, lambda: self.stop_btn.configure(state="disabled"))
            self.dialog.after(0, lambda: messagebox.showerror("API 错误", msg))
        except:
            pass
        self.is_generating = False
    
    def _stop(self):
        self.is_generating = False
        self.status_label.configure(text="⏹ 已停止")
        self.stop_btn.configure(state="disabled")
    
    def _copy(self):
        result = self.output_text.get("1.0", "end-1c")
        if result:
            self.dialog.clipboard_clear()
            self.dialog.clipboard_append(result)
            self.status_label.configure(text="📋 已复制")
    
    def _insert(self):
        result = self.output_text.get("1.0", "end-1c")
        if not result:
            return
        textbox = self._get_textbox()
        if textbox:
            textbox.insert("insert", result)
            self.dialog.destroy()
            self.dialog = None
    
    def _replace(self):
        result = self.output_text.get("1.0", "end-1c")
        if not result:
            return
        textbox = self._get_textbox()
        if textbox:
            try:
                textbox.delete("sel.first", "sel.last")
                textbox.insert("insert", result)
                self.dialog.destroy()
                self.dialog = None
            except:
                messagebox.showinfo("提示", "没有选中文本，将在光标处插入")
                textbox.insert("insert", result)
    
    # 兼容旧接口
    def show_ai_dialog(self):
        self.show_dialog()
