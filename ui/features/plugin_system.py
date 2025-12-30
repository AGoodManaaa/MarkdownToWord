# -*- coding: utf-8 -*-
"""
插件系统 - 支持第三方插件扩展
"""

import os
import json
import importlib.util
from typing import Dict, List, Callable
import customtkinter as ctk
from tkinter import filedialog, messagebox


class Plugin:
    """插件基类"""
    
    def __init__(self):
        self.name = "未命名插件"
        self.version = "1.0.0"
        self.author = "未知"
        self.description = "无描述"
        
    def activate(self, app):
        """激活插件"""
        pass
        
    def deactivate(self, app):
        """停用插件"""
        pass
        
    def get_menu_items(self) -> List[tuple]:
        """
        返回插件菜单项
        返回格式: [(name, callback), ...]
        """
        return []
        
    def on_document_open(self, content: str) -> str:
        """文档打开时的钩子"""
        return content
        
    def on_document_save(self, content: str) -> str:
        """文档保存时的钩子"""
        return content
        
    def on_export(self, content: str) -> str:
        """导出前的钩子"""
        return content


class PluginManager:
    """插件管理器"""
    
    def __init__(self, app):
        self.app = app
        self.plugins: Dict[str, Plugin] = {}
        self.active_plugins: set = set()
        self.plugins_dir = self._get_plugins_dir()
        self.manager_dialog = None
        
        # 加载已安装的插件
        self._load_plugins()
        
    def _get_plugins_dir(self) -> str:
        """获取插件目录"""
        app_dir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
        plugins_dir = os.path.join(app_dir, "plugins")
        os.makedirs(plugins_dir, exist_ok=True)
        return plugins_dir
        
    def _load_plugins(self):
        """加载所有插件"""
        if not os.path.exists(self.plugins_dir):
            return
            
        for item in os.listdir(self.plugins_dir):
            plugin_path = os.path.join(self.plugins_dir, item)
            
            # 检查是否是有效的插件目录
            if os.path.isdir(plugin_path):
                manifest_path = os.path.join(plugin_path, "plugin.json")
                main_path = os.path.join(plugin_path, "main.py")
                
                if os.path.exists(manifest_path) and os.path.exists(main_path):
                    try:
                        self._load_plugin(plugin_path, manifest_path, main_path)
                    except Exception as e:
                        print(f"加载插件失败 {item}: {e}")
                        
    def _load_plugin(self, plugin_dir: str, manifest_path: str, main_path: str):
        """加载单个插件"""
        # 读取清单
        with open(manifest_path, 'r', encoding='utf-8') as f:
            manifest = json.load(f)
            
        plugin_id = manifest.get('id', os.path.basename(plugin_dir))
        
        # 动态加载Python模块
        spec = importlib.util.spec_from_file_location(plugin_id, main_path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        
        # 创建插件实例
        if hasattr(module, 'PluginClass'):
            plugin = module.PluginClass()
            plugin.name = manifest.get('name', plugin_id)
            plugin.version = manifest.get('version', '1.0.0')
            plugin.author = manifest.get('author', '未知')
            plugin.description = manifest.get('description', '')
            
            self.plugins[plugin_id] = plugin
            
    def show_plugin_manager(self):
        """显示插件管理器"""
        if self.manager_dialog and self.manager_dialog.winfo_exists():
            self.manager_dialog.focus()
            return
            
        self.manager_dialog = ctk.CTkToplevel(self.app)
        self.manager_dialog.title("🔌 插件管理器")
        self.manager_dialog.geometry("800x600")
        self.manager_dialog.transient(self.app)
        
        # 标题
        title_label = ctk.CTkLabel(
            self.manager_dialog,
            text="插件管理器",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        title_label.pack(pady=15)
        
        # 选项卡
        tabview = ctk.CTkTabview(self.manager_dialog, width=760, height=480)
        tabview.pack(padx=20, pady=10)
        
        tabview.add("已安装")
        tabview.add("插件市场")
        tabview.add("开发")
        
        # === 已安装标签页 ===
        self._create_installed_tab(tabview.tab("已安装"))
        
        # === 插件市场标签页 ===
        self._create_marketplace_tab(tabview.tab("插件市场"))
        
        # === 开发标签页 ===
        self._create_development_tab(tabview.tab("开发"))
        
        # 底部按钮
        btn_frame = ctk.CTkFrame(self.manager_dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkButton(
            btn_frame,
            text="📁 打开插件目录",
            command=self._open_plugins_dir,
            width=140
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            btn_frame,
            text="关闭",
            command=self.manager_dialog.destroy,
            fg_color="#6B7280",
            hover_color="#4B5563",
            width=100
        ).pack(side="right", padx=5)
        
    def _create_installed_tab(self, parent):
        """创建已安装插件标签页"""
        if not self.plugins:
            ctk.CTkLabel(
                parent,
                text="暂无已安装的插件",
                text_color="#9CA3AF",
                font=ctk.CTkFont(size=14)
            ).pack(pady=50)
            
            install_btn = ctk.CTkButton(
                parent,
                text="📥 安装插件",
                command=self._install_plugin,
                fg_color="#10B981",
                hover_color="#059669"
            )
            install_btn.pack(pady=10)
            return
            
        # 插件列表
        list_frame = ctk.CTkScrollableFrame(parent, width=720, height=420)
        list_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        for plugin_id, plugin in self.plugins.items():
            self._create_plugin_card(list_frame, plugin_id, plugin)
            
    def _create_plugin_card(self, parent, plugin_id: str, plugin: Plugin):
        """创建插件卡片"""
        card = ctk.CTkFrame(parent, fg_color="#F9FAFB", corner_radius=10)
        card.pack(fill="x", pady=5, padx=5)
        
        # 左侧信息
        info_frame = ctk.CTkFrame(card, fg_color="transparent")
        info_frame.pack(side="left", fill="both", expand=True, padx=15, pady=10)
        
        # 插件名称
        name_label = ctk.CTkLabel(
            info_frame,
            text=plugin.name,
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w"
        )
        name_label.pack(anchor="w", pady=2)
        
        # 版本和作者
        meta_label = ctk.CTkLabel(
            info_frame,
            text=f"v{plugin.version} · {plugin.author}",
            font=ctk.CTkFont(size=11),
            text_color="#6B7280",
            anchor="w"
        )
        meta_label.pack(anchor="w", pady=2)
        
        # 描述
        desc_label = ctk.CTkLabel(
            info_frame,
            text=plugin.description,
            font=ctk.CTkFont(size=12),
            text_color="#4B5563",
            anchor="w",
            wraplength=400
        )
        desc_label.pack(anchor="w", pady=5)
        
        # 右侧按钮
        btn_frame = ctk.CTkFrame(card, fg_color="transparent")
        btn_frame.pack(side="right", padx=15, pady=10)
        
        is_active = plugin_id in self.active_plugins
        
        toggle_btn = ctk.CTkButton(
            btn_frame,
            text="✅ 已启用" if is_active else "启用",
            command=lambda: self._toggle_plugin(plugin_id, toggle_btn),
            width=100,
            fg_color="#10B981" if is_active else "#6B7280",
            hover_color="#059669" if is_active else "#4B5563"
        )
        toggle_btn.pack(pady=2)
        
        settings_btn = ctk.CTkButton(
            btn_frame,
            text="⚙️",
            command=lambda: self._show_plugin_settings(plugin_id),
            width=50
        )
        settings_btn.pack(pady=2)
        
        delete_btn = ctk.CTkButton(
            btn_frame,
            text="🗑️",
            command=lambda: self._uninstall_plugin(plugin_id),
            width=50,
            fg_color="#EF4444",
            hover_color="#DC2626"
        )
        delete_btn.pack(pady=2)
        
    def _create_marketplace_tab(self, parent):
        """创建插件市场标签页"""
        ctk.CTkLabel(
            parent,
            text="插件市场",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=15)
        
        # 示例插件列表
        marketplace_plugins = [
            {
                "name": "Markdown增强",
                "author": "官方",
                "description": "提供更多Markdown语法支持",
                "downloads": 1234,
                "rating": 4.8
            },
            {
                "name": "代码美化",
                "author": "社区",
                "description": "自动美化代码块格式",
                "downloads": 856,
                "rating": 4.5
            },
            {
                "name": "图床上传",
                "author": "社区",
                "description": "支持图片自动上传到图床",
                "downloads": 2341,
                "rating": 4.9
            }
        ]
        
        list_frame = ctk.CTkScrollableFrame(parent, width=720, height=380)
        list_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        for plugin_info in marketplace_plugins:
            self._create_marketplace_card(list_frame, plugin_info)
            
    def _create_marketplace_card(self, parent, info: dict):
        """创建市场插件卡片"""
        card = ctk.CTkFrame(parent, fg_color="#F9FAFB", corner_radius=10)
        card.pack(fill="x", pady=5, padx=5)
        
        # 信息
        info_frame = ctk.CTkFrame(card, fg_color="transparent")
        info_frame.pack(side="left", fill="both", expand=True, padx=15, pady=10)
        
        ctk.CTkLabel(
            info_frame,
            text=info['name'],
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w"
        ).pack(anchor="w", pady=2)
        
        ctk.CTkLabel(
            info_frame,
            text=f"👤 {info['author']} · 📥 {info['downloads']} · ⭐ {info['rating']}",
            font=ctk.CTkFont(size=11),
            text_color="#6B7280",
            anchor="w"
        ).pack(anchor="w", pady=2)
        
        ctk.CTkLabel(
            info_frame,
            text=info['description'],
            font=ctk.CTkFont(size=12),
            text_color="#4B5563",
            anchor="w"
        ).pack(anchor="w", pady=5)
        
        # 安装按钮
        ctk.CTkButton(
            card,
            text="📥 安装",
            command=lambda: messagebox.showinfo("提示", "插件市场功能开发中..."),
            width=100,
            fg_color="#10B981",
            hover_color="#059669"
        ).pack(side="right", padx=15, pady=10)
        
    def _create_development_tab(self, parent):
        """创建开发标签页"""
        ctk.CTkLabel(
            parent,
            text="插件开发指南",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=15)
        
        guide_text = ctk.CTkTextbox(parent, width=720, height=420)
        guide_text.pack(padx=10, pady=10)
        
        guide_text.insert("1.0", """
# 插件开发指南

## 插件结构

每个插件需要包含以下文件：

```
my-plugin/
├── plugin.json    # 插件清单
└── main.py        # 插件主文件
```

## plugin.json 示例

```json
{
  "id": "my-plugin",
  "name": "我的插件",
  "version": "1.0.0",
  "author": "作者名",
  "description": "插件描述",
  "homepage": "https://example.com"
}
```

## main.py 示例

```python
from ui.features.plugin_system import Plugin

class PluginClass(Plugin):
    def __init__(self):
        super().__init__()
        self.name = "我的插件"
        
    def activate(self, app):
        '''插件激活时调用'''
        print("插件已激活")
        
    def on_export(self, content):
        '''导出前处理内容'''
        # 在这里处理Markdown内容
        return content
```

## 可用钩子

- `activate(app)` - 插件激活
- `deactivate(app)` - 插件停用
- `on_document_open(content)` - 文档打开
- `on_document_save(content)` - 文档保存
- `on_export(content)` - 导出前处理

## API 文档

更多API文档请访问：https://github.com/AGoodManaaa/MarkdowntoWord/wiki
""")
        guide_text.configure(state="disabled")
        
    def _toggle_plugin(self, plugin_id: str, button):
        """启用/停用插件"""
        if plugin_id in self.active_plugins:
            # 停用
            self.active_plugins.remove(plugin_id)
            self.plugins[plugin_id].deactivate(self.app)
            button.configure(text="启用", fg_color="#6B7280", hover_color="#4B5563")
            self.app.update_status(f"❌ 已停用插件: {self.plugins[plugin_id].name}")
        else:
            # 启用
            self.active_plugins.add(plugin_id)
            self.plugins[plugin_id].activate(self.app)
            button.configure(text="✅ 已启用", fg_color="#10B981", hover_color="#059669")
            self.app.update_status(f"✅ 已启用插件: {self.plugins[plugin_id].name}")
            
    def _show_plugin_settings(self, plugin_id: str):
        """显示插件设置"""
        messagebox.showinfo("插件设置", f"{self.plugins[plugin_id].name}\n\n插件设置功能开发中...")
        
    def _uninstall_plugin(self, plugin_id: str):
        """卸载插件"""
        result = messagebox.askyesno(
            "确认卸载",
            f"确定要卸载插件 {self.plugins[plugin_id].name} 吗？"
        )
        
        if result:
            # 停用插件
            if plugin_id in self.active_plugins:
                self.active_plugins.remove(plugin_id)
                self.plugins[plugin_id].deactivate(self.app)
                
            # 删除插件文件
            import shutil
            plugin_dir = os.path.join(self.plugins_dir, plugin_id)
            if os.path.exists(plugin_dir):
                shutil.rmtree(plugin_dir)
                
            # 从列表移除
            del self.plugins[plugin_id]
            
            messagebox.showinfo("完成", "插件已卸载")
            
            # 刷新界面
            if self.manager_dialog:
                self.manager_dialog.destroy()
                self.show_plugin_manager()
                
    def _install_plugin(self):
        """安装插件"""
        # 选择插件文件（.zip）
        file_path = filedialog.askopenfilename(
            title="选择插件文件",
            filetypes=[("插件包", "*.zip"), ("所有文件", "*.*")]
        )
        
        if file_path:
            messagebox.showinfo("提示", "插件安装功能开发中...")
            
    def _open_plugins_dir(self):
        """打开插件目录"""
        import subprocess
        import platform
        
        if platform.system() == 'Windows':
            os.startfile(self.plugins_dir)
        elif platform.system() == 'Darwin':
            subprocess.run(['open', self.plugins_dir])
        else:
            subprocess.run(['xdg-open', self.plugins_dir])
            
    def call_hook(self, hook_name: str, *args, **kwargs):
        """调用所有活动插件的钩子"""
        result = args[0] if args else None
        
        for plugin_id in self.active_plugins:
            plugin = self.plugins[plugin_id]
            if hasattr(plugin, hook_name):
                try:
                    method = getattr(plugin, hook_name)
                    result = method(*args, **kwargs)
                except Exception as e:
                    print(f"插件 {plugin.name} 钩子 {hook_name} 执行失败: {e}")
                    
        return result
