# -*- coding: utf-8 -*-
"""Markdown 语法自动补全功能"""

import tkinter as tk
import customtkinter as ctk

class AutocompleteFeature:
    """Markdown 语法自动补全
    
    支持:
    - 输入 '#' 自动补全标题
    - 输入 '[' 自动补全链接
    - 输入 '```' 自动补全代码块
    - 输入 '>' 自动补全引用
    - 输入 '-' 自动补全列表
    """
    
    def __init__(self, app):
        self.app = app
        self._popup = None
        self._bind_events()
        
    def _bind_events(self):
        """绑定编辑器事件"""
        if hasattr(self.app, 'input_text'):
            textbox = self.app.input_text._textbox
            textbox.bind('<Key>', self._on_key)
            
    def _on_key(self, event):
        """按键事件处理"""
        if event.char in ['#', '[', '`', '>', '-', '!', '(']:
            self.app.after(10, lambda: self._check_trigger(event.char))
            
    def _check_trigger(self, char):
        """检查触发词"""
        try:
            textbox = self.app.input_text._textbox
            index = textbox.index(tk.INSERT)
            line_start = f"{index.split('.')[0]}.0"
            line_text = textbox.get(line_start, index)
            
            # 自动补全标题
            if char == '#' and line_text.strip() == '#':
                self._show_suggestions(index, ['# ', '## ', '### ', '#### ', '##### ', '###### '])
                
            # 自动补全代码块
            elif char == '`' and line_text.endswith('```'):
                self._show_suggestions(index, [
                    '```python\n\n```',
                    '```javascript\n\n```', 
                    '```html\n\n```',
                    '```css\n\n```',
                    '```bash\n\n```',
                    '```sql\n\n```'
                ])
                
            # 自动补全链接/图片
            elif char == '[':
                self._show_suggestions(index, ['[链接文字](url)', '![图片说明](url)'])
            
            # 自动补全引用
            elif char == '>' and line_text.strip() == '>':
                self._show_suggestions(index, ['> '])
                
            # 自动补全列表
            elif char == '-' and line_text.strip() == '-':
                self._show_suggestions(index, ['- ', '- [ ] '])
                
        except Exception:
            pass
            
    def _show_suggestions(self, index, suggestions):
        """显示补全建议弹窗"""
        if self._popup:
            self._popup.destroy()
            
        textbox = self.app.input_text._textbox
        bbox = textbox.bbox(index)
        if not bbox:
            return
            
        x, y, _, h = bbox
        root_x = textbox.winfo_rootx() + x
        root_y = textbox.winfo_rooty() + y + h
        
        self._popup = ctk.CTkToplevel(self.app)
        self._popup.overrideredirect(True)
        self._popup.geometry(f"+{root_x}+{root_y}")
        self._popup.attributes('-topmost', True)
        
        frame = ctk.CTkFrame(self._popup, fg_color="gray90", corner_radius=5)
        frame.pack(fill='both', expand=True, padx=1, pady=1)
        
        for text in suggestions:
            btn = ctk.CTkButton(
                frame,
                text=text.strip() or "Empty",
                anchor='w',
                fg_color='transparent',
                text_color='black',
                hover_color='gray80',
                height=25,
                command=lambda t=text: self._apply_completion(t)
            )
            btn.pack(fill='x')
            
        # 点击其他地方关闭
        self._popup.bind('<FocusOut>', lambda e: self._popup.destroy())
        self._popup.focus_set()
        
    def _apply_completion(self, text):
        """应用补全"""
        textbox = self.app.input_text._textbox
        
        # 删除触发字符/文本
        if text.startswith('```'):
            # 对于代码块，不仅仅是插入，可能需要删除之前的 ```
            # 这里简化处理：假设用户刚输入了 ```
            pass 
            
        # 简单插入
        textbox.insert(tk.INSERT, text)
        
        if self._popup:
            self._popup.destroy()
            self._popup = None
