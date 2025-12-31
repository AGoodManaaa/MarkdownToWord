# -*- coding: utf-8 -*-
"""
快捷键自定义功能
支持查看、修改、重置快捷键
"""

import json
import os
import customtkinter as ctk
from tkinter import messagebox
from typing import Dict, Callable, Optional, List, Tuple
from dataclasses import dataclass, field, asdict
from ui.theme import COLORS
from ui.dialog_utils import set_dialog_icon


@dataclass
class ShortcutItem:
    """快捷键项"""
    id: str  # 唯一标识
    name: str  # 显示名称
    description: str  # 描述
    default_key: str  # 默认快捷键
    current_key: str  # 当前快捷键
    category: str = "通用"  # 分类


class KeyboardShortcutsManager:
    """快捷键管理器"""
    
    def __init__(self, app):
        self.app = app
        self.shortcuts: Dict[str, ShortcutItem] = {}
        self.callbacks: Dict[str, Callable] = {}
        self._config_file = os.path.join(os.path.dirname(__file__), '..', '..', 'shortcuts.json')
        
        # 初始化默认快捷键
        self._init_default_shortcuts()
        # 加载用户配置
        self._load_config()
    
    def _init_default_shortcuts(self):
        """初始化默认快捷键"""
        defaults = [
            # 文件操作
            ("file_new", "新建文件", "创建新文档", "<Control-n>", "文件"),
            ("file_open", "打开文件", "打开现有文档", "<Control-o>", "文件"),
            ("file_save", "保存文件", "保存当前文档", "<Control-s>", "文件"),
            ("file_save_as", "另存为", "将文档另存为新文件", "<Control-Shift-s>", "文件"),
            
            # 编辑操作
            ("edit_undo", "撤销", "撤销上一步操作", "<Control-z>", "编辑"),
            ("edit_redo", "重做", "重做上一步操作", "<Control-y>", "编辑"),
            ("edit_cut", "剪切", "剪切选中内容", "<Control-x>", "编辑"),
            ("edit_copy", "复制", "复制选中内容", "<Control-c>", "编辑"),
            ("edit_paste", "粘贴", "粘贴剪贴板内容", "<Control-v>", "编辑"),
            ("edit_select_all", "全选", "选择全部内容", "<Control-a>", "编辑"),
            
            # 搜索
            ("search_find", "查找", "打开查找对话框", "<Control-f>", "搜索"),
            ("search_replace", "替换", "打开替换对话框", "<Control-h>", "搜索"),
            ("search_next", "查找下一个", "跳转到下一个匹配", "<F3>", "搜索"),
            ("search_prev", "查找上一个", "跳转到上一个匹配", "<Shift-F3>", "搜索"),
            
            # 视图
            ("view_preview", "切换预览", "显示/隐藏预览面板", "<Control-p>", "视图"),
            ("view_sidebar", "切换侧边栏", "显示/隐藏侧边栏", "<Control-b>", "视图"),
            ("view_minimap", "切换迷你地图", "显示/隐藏迷你地图", "<Control-m>", "视图"),
            ("view_fullscreen", "全屏预览", "进入全屏预览模式", "<Control-F11>", "视图"),
            ("view_focus", "专注模式", "进入专注模式", "<F11>", "视图"),
            ("view_reading", "阅读模式", "进入阅读模式", "<F12>", "视图"),
            
            # 字体
            ("font_increase", "增大字体", "增大编辑器字体", "<Control-plus>", "字体"),
            ("font_decrease", "减小字体", "减小编辑器字体", "<Control-minus>", "字体"),
            
            # 导出
            ("export_word", "导出Word", "导出为Word文档", "<Control-e>", "导出"),
            ("export_pdf", "导出PDF", "导出为PDF文档", "<Control-Shift-e>", "导出"),
            ("export_html", "导出HTML", "导出为HTML文件", "<Control-Alt-e>", "导出"),
            
            # 工具
            ("tool_format", "格式化", "格式化Markdown文档", "<Control-Shift-f>", "工具"),
            ("tool_toc", "插入目录", "插入文档目录", "<Control-Shift-t>", "工具"),
            ("tool_command", "命令面板", "打开命令面板", "<Control-k>", "工具"),
            ("tool_help", "帮助", "显示帮助信息", "<F1>", "工具"),
        ]
        
        for id, name, desc, key, cat in defaults:
            self.shortcuts[id] = ShortcutItem(
                id=id, name=name, description=desc,
                default_key=key, current_key=key, category=cat
            )
    
    def _load_config(self):
        """加载用户配置"""
        try:
            if os.path.exists(self._config_file):
                with open(self._config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    for id, key in data.items():
                        if id in self.shortcuts:
                            self.shortcuts[id].current_key = key
        except Exception as e:
            print(f"加载快捷键配置失败: {e}")
    
    def _save_config(self):
        """保存用户配置"""
        try:
            data = {id: s.current_key for id, s in self.shortcuts.items()}
            with open(self._config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存快捷键配置失败: {e}")
    
    def register(self, shortcut_id: str, callback: Callable):
        """注册快捷键回调"""
        self.callbacks[shortcut_id] = callback
        if shortcut_id in self.shortcuts:
            key = self.shortcuts[shortcut_id].current_key
            try:
                self.app.bind(key, lambda e, cb=callback: cb())
            except Exception:
                pass
    
    def update_shortcut(self, shortcut_id: str, new_key: str) -> Tuple[bool, str]:
        """更新快捷键"""
        if shortcut_id not in self.shortcuts:
            return False, "快捷键ID不存在"
        
        # 检查冲突
        for id, s in self.shortcuts.items():
            if id != shortcut_id and s.current_key == new_key:
                return False, f"快捷键已被 '{s.name}' 使用"
        
        old_key = self.shortcuts[shortcut_id].current_key
        self.shortcuts[shortcut_id].current_key = new_key
        
        # 重新绑定
        if shortcut_id in self.callbacks:
            try:
                self.app.unbind(old_key)
            except Exception:
                pass
            try:
                self.app.bind(new_key, lambda e, cb=self.callbacks[shortcut_id]: cb())
            except Exception:
                pass
        
        self._save_config()
        return True, "快捷键已更新"
    
    def reset_shortcut(self, shortcut_id: str):
        """重置单个快捷键"""
        if shortcut_id in self.shortcuts:
            default = self.shortcuts[shortcut_id].default_key
            self.update_shortcut(shortcut_id, default)
    
    def reset_all(self):
        """重置所有快捷键"""
        for id, s in self.shortcuts.items():
            s.current_key = s.default_key
        self._save_config()
        self._rebind_all()
    
    def _rebind_all(self):
        """重新绑定所有快捷键"""
        for id, callback in self.callbacks.items():
            if id in self.shortcuts:
                key = self.shortcuts[id].current_key
                try:
                    self.app.bind(key, lambda e, cb=callback: cb())
                except Exception:
                    pass
    
    def get_by_category(self) -> Dict[str, List[ShortcutItem]]:
        """按分类获取快捷键"""
        result: Dict[str, List[ShortcutItem]] = {}
        for s in self.shortcuts.values():
            if s.category not in result:
                result[s.category] = []
            result[s.category].append(s)
        return result


class KeyboardShortcutsDialog:
    """快捷键设置对话框"""
    
    def __init__(self, app, manager: KeyboardShortcutsManager):
        self.app = app
        self.manager = manager
        self.dialog = None
        self.editing_id = None
        self.key_entries: Dict[str, ctk.CTkEntry] = {}
    
    def show(self):
        """显示对话框"""
        if self.dialog is not None and self.dialog.winfo_exists():
            self.dialog.focus()
            return
        
        self.dialog = ctk.CTkToplevel(self.app)
        self.dialog.title("⌨️ 快捷键设置")
        self.dialog.geometry("700x550")
        self.dialog.transient(self.app)
        set_dialog_icon(self.dialog)
        
        # 居中
        self.dialog.update_idletasks()
        x = self.app.winfo_x() + (self.app.winfo_width() - 700) // 2
        y = self.app.winfo_y() + (self.app.winfo_height() - 550) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        self._create_ui()
    
    def _create_ui(self):
        """创建界面"""
        # 标题
        header = ctk.CTkFrame(self.dialog, fg_color="transparent")
        header.pack(fill="x", padx=20, pady=(15, 10))
        
        ctk.CTkLabel(
            header, text="⌨️ 快捷键设置",
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(side="left")
        
        # 重置按钮
        ctk.CTkButton(
            header, text="🔄 重置全部", width=100,
            fg_color="gray", hover_color="#666666",
            command=self._reset_all
        ).pack(side="right")
        
        # 搜索框
        search_frame = ctk.CTkFrame(self.dialog, fg_color="transparent")
        search_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        self.search_var = ctk.StringVar()
        self.search_var.trace_add("write", self._on_search)
        
        ctk.CTkEntry(
            search_frame, textvariable=self.search_var,
            placeholder_text="🔍 搜索快捷键...", width=300
        ).pack(side="left")
        
        # 快捷键列表（滚动区域）
        self.list_frame = ctk.CTkScrollableFrame(self.dialog, fg_color="transparent")
        self.list_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        
        self._populate_list()
        
        # 底部提示
        tip = ctk.CTkLabel(
            self.dialog,
            text="💡 点击快捷键输入框，然后按下新的快捷键组合",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        )
        tip.pack(pady=(0, 15))
    
    def _populate_list(self, filter_text: str = ""):
        """填充快捷键列表"""
        # 清空
        for widget in self.list_frame.winfo_children():
            widget.destroy()
        self.key_entries.clear()
        
        categories = self.manager.get_by_category()
        filter_lower = filter_text.lower()
        
        for cat_name, shortcuts in categories.items():
            # 过滤
            filtered = [s for s in shortcuts if 
                       filter_lower in s.name.lower() or 
                       filter_lower in s.description.lower() or
                       filter_lower in s.current_key.lower()]
            
            if not filtered:
                continue
            
            # 分类标题
            cat_frame = ctk.CTkFrame(self.list_frame, fg_color=COLORS['bg_sidebar'])
            cat_frame.pack(fill="x", pady=(10, 5))
            
            ctk.CTkLabel(
                cat_frame, text=f"📁 {cat_name}",
                font=ctk.CTkFont(size=13, weight="bold"),
                text_color=COLORS['primary']
            ).pack(anchor="w", padx=10, pady=5)
            
            # 快捷键项
            for shortcut in filtered:
                self._create_shortcut_row(shortcut)
    
    def _create_shortcut_row(self, shortcut: ShortcutItem):
        """创建快捷键行"""
        row = ctk.CTkFrame(self.list_frame, fg_color=COLORS['bg_card'], corner_radius=8)
        row.pack(fill="x", pady=2, padx=5)
        
        # 左侧：名称和描述
        left = ctk.CTkFrame(row, fg_color="transparent")
        left.pack(side="left", fill="x", expand=True, padx=10, pady=8)
        
        ctk.CTkLabel(
            left, text=shortcut.name,
            font=ctk.CTkFont(size=12, weight="bold"),
            anchor="w"
        ).pack(anchor="w")
        
        ctk.CTkLabel(
            left, text=shortcut.description,
            font=ctk.CTkFont(size=10),
            text_color="gray", anchor="w"
        ).pack(anchor="w")
        
        # 右侧：快捷键输入
        right = ctk.CTkFrame(row, fg_color="transparent")
        right.pack(side="right", padx=10, pady=8)
        
        # 显示格式化的快捷键
        display_key = self._format_key(shortcut.current_key)
        
        key_entry = ctk.CTkEntry(
            right, width=150, justify="center",
            font=ctk.CTkFont(size=11)
        )
        key_entry.insert(0, display_key)
        key_entry.pack(side="left", padx=(0, 5))
        
        # 绑定按键捕获
        key_entry.bind('<FocusIn>', lambda e, sid=shortcut.id: self._start_capture(sid, e.widget))
        key_entry.bind('<KeyPress>', lambda e, sid=shortcut.id: self._on_key_press(sid, e))
        
        self.key_entries[shortcut.id] = key_entry
        
        # 重置按钮
        if shortcut.current_key != shortcut.default_key:
            ctk.CTkButton(
                right, text="↩", width=30, height=28,
                fg_color="gray", hover_color="#666666",
                command=lambda sid=shortcut.id: self._reset_one(sid)
            ).pack(side="left")
    
    def _format_key(self, key: str) -> str:
        """格式化快捷键显示"""
        # <Control-Shift-s> -> Ctrl+Shift+S
        key = key.replace('<', '').replace('>', '')
        key = key.replace('Control', 'Ctrl')
        key = key.replace('-', '+')
        parts = key.split('+')
        parts = [p.upper() if len(p) == 1 else p for p in parts]
        return '+'.join(parts)
    
    def _parse_key(self, display: str) -> str:
        """解析显示格式为绑定格式"""
        # Ctrl+Shift+S -> <Control-Shift-s>
        display = display.replace('Ctrl', 'Control')
        display = display.replace('+', '-')
        parts = display.split('-')
        parts = [p.lower() if len(p) == 1 else p for p in parts]
        return f"<{'-'.join(parts)}>"
    
    def _start_capture(self, shortcut_id: str, widget):
        """开始捕获按键"""
        self.editing_id = shortcut_id
        widget.delete(0, 'end')
        widget.insert(0, "按下快捷键...")
        widget.configure(fg_color=COLORS['highlight'])
    
    def _on_key_press(self, shortcut_id: str, event):
        """按键事件"""
        if self.editing_id != shortcut_id:
            return
        
        # 忽略单独的修饰键
        if event.keysym in ('Control_L', 'Control_R', 'Shift_L', 'Shift_R', 
                           'Alt_L', 'Alt_R', 'Super_L', 'Super_R'):
            return "break"
        
        # 构建快捷键
        parts = []
        if event.state & 0x4:  # Control
            parts.append('Control')
        if event.state & 0x1:  # Shift
            parts.append('Shift')
        if event.state & 0x8:  # Alt
            parts.append('Alt')
        
        # 添加按键
        key = event.keysym
        if len(key) == 1:
            key = key.lower()
        parts.append(key)
        
        new_key = f"<{'-'.join(parts)}>"
        
        # 更新
        success, msg = self.manager.update_shortcut(shortcut_id, new_key)
        
        entry = self.key_entries.get(shortcut_id)
        if entry:
            entry.delete(0, 'end')
            if success:
                entry.insert(0, self._format_key(new_key))
                entry.configure(fg_color=COLORS['bg_card'])
            else:
                entry.insert(0, msg)
                entry.configure(fg_color="#ffcccc")
                self.dialog.after(1500, lambda: self._restore_entry(shortcut_id))
        
        self.editing_id = None
        return "break"
    
    def _restore_entry(self, shortcut_id: str):
        """恢复输入框"""
        if shortcut_id in self.manager.shortcuts:
            entry = self.key_entries.get(shortcut_id)
            if entry:
                entry.delete(0, 'end')
                entry.insert(0, self._format_key(self.manager.shortcuts[shortcut_id].current_key))
                entry.configure(fg_color=COLORS['bg_card'])
    
    def _reset_one(self, shortcut_id: str):
        """重置单个快捷键"""
        self.manager.reset_shortcut(shortcut_id)
        self._populate_list(self.search_var.get())
    
    def _reset_all(self):
        """重置所有快捷键"""
        if messagebox.askyesno("确认", "确定要重置所有快捷键为默认值吗？", parent=self.dialog):
            self.manager.reset_all()
            self._populate_list(self.search_var.get())
            messagebox.showinfo("完成", "所有快捷键已重置", parent=self.dialog)
    
    def _on_search(self, *args):
        """搜索过滤"""
        self._populate_list(self.search_var.get())


class KeyboardShortcutsFeature:
    """快捷键自定义功能"""
    
    def __init__(self, app):
        self.app = app
        self.manager = KeyboardShortcutsManager(app)
        self.dialog = None
    
    def show_dialog(self):
        """显示快捷键设置对话框"""
        if self.dialog is None:
            self.dialog = KeyboardShortcutsDialog(self.app, self.manager)
        self.dialog.show()
    
    def register_shortcuts(self):
        """注册所有快捷键"""
        # 这里可以注册应用的快捷键回调
        pass
