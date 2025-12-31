# -*- coding: utf-8 -*-
"""
文件夹视图功能
在侧边栏显示文件树，支持浏览、打开、右键菜单
"""

import os
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox
from typing import Optional, Callable, Dict, List
from dataclasses import dataclass
from ui.theme import COLORS


@dataclass
class FileNode:
    """文件节点"""
    path: str
    name: str
    is_dir: bool
    parent: Optional['FileNode'] = None
    children: List['FileNode'] = None
    expanded: bool = False
    
    def __post_init__(self):
        if self.children is None:
            self.children = []


class FolderTreeView(ctk.CTkFrame):
    """文件夹树视图组件"""
    
    def __init__(self, parent, on_file_open: Callable[[str], None] = None, **kwargs):
        super().__init__(parent, **kwargs)
        
        self.on_file_open = on_file_open
        self.root_path: Optional[str] = None
        self.root_node: Optional[FileNode] = None
        self.selected_path: Optional[str] = None
        self._item_widgets: Dict[str, ctk.CTkFrame] = {}
        self._expanded_paths: set = set()
        
        # 文件图标
        self._icons = {
            'folder': '📁',
            'folder_open': '📂',
            'md': '📝',
            'txt': '📄',
            'py': '🐍',
            'json': '📋',
            'html': '🌐',
            'css': '🎨',
            'js': '⚡',
            'image': '🖼️',
            'default': '📄',
        }
        
        self._create_ui()
    
    def _create_ui(self):
        """创建界面"""
        # 头部
        header = ctk.CTkFrame(self, fg_color="transparent", height=35)
        header.pack(fill="x", padx=5, pady=(5, 0))
        header.pack_propagate(False)
        
        self.title_label = ctk.CTkLabel(
            header, text="📁 文件夹",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=COLORS['text_secondary']
        )
        self.title_label.pack(side="left", padx=5)
        
        # 打开文件夹按钮
        self.open_btn = ctk.CTkButton(
            header, text="📂", width=28, height=24,
            fg_color="transparent", hover_color=COLORS['highlight'],
            text_color=COLORS['text_primary'],
            command=self.open_folder
        )
        self.open_btn.pack(side="right", padx=2)
        
        # 刷新按钮
        self.refresh_btn = ctk.CTkButton(
            header, text="🔄", width=28, height=24,
            fg_color="transparent", hover_color=COLORS['highlight'],
            text_color=COLORS['text_primary'],
            command=self.refresh
        )
        self.refresh_btn.pack(side="right", padx=2)
        
        # 折叠全部按钮
        self.collapse_btn = ctk.CTkButton(
            header, text="⊟", width=28, height=24,
            fg_color="transparent", hover_color=COLORS['highlight'],
            text_color=COLORS['text_primary'],
            command=self.collapse_all
        )
        self.collapse_btn.pack(side="right", padx=2)
        
        # 树视图容器（滚动）
        self.tree_container = ctk.CTkScrollableFrame(
            self, fg_color="transparent",
            scrollbar_button_color=COLORS['border'],
            scrollbar_button_hover_color=COLORS['primary']
        )
        self.tree_container.pack(fill="both", expand=True, padx=2, pady=5)
        
        # 空状态提示
        self.empty_label = ctk.CTkLabel(
            self.tree_container,
            text="点击 📂 打开文件夹",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        )
        self.empty_label.pack(pady=20)
    
    def open_folder(self, path: str = None):
        """打开文件夹"""
        if path is None:
            path = filedialog.askdirectory(title="选择文件夹")
        
        if path and os.path.isdir(path):
            self.root_path = path
            self._expanded_paths.clear()
            self._expanded_paths.add(path)  # 根目录默认展开
            self.refresh()
    
    def refresh(self):
        """刷新文件树"""
        if not self.root_path or not os.path.exists(self.root_path):
            return
        
        # 清空
        for widget in self.tree_container.winfo_children():
            widget.destroy()
        self._item_widgets.clear()
        
        # 构建树
        self.root_node = self._build_tree(self.root_path)
        
        # 更新标题
        folder_name = os.path.basename(self.root_path)
        self.title_label.configure(text=f"📁 {folder_name}")
        
        # 渲染树
        self._render_tree(self.root_node, 0)
    
    def _build_tree(self, path: str, parent: FileNode = None) -> FileNode:
        """构建文件树"""
        name = os.path.basename(path) or path
        is_dir = os.path.isdir(path)
        node = FileNode(path=path, name=name, is_dir=is_dir, parent=parent)
        
        if is_dir:
            try:
                entries = sorted(os.listdir(path), key=lambda x: (not os.path.isdir(os.path.join(path, x)), x.lower()))
                for entry in entries:
                    # 跳过隐藏文件和特殊目录
                    if entry.startswith('.') or entry in ('__pycache__', 'node_modules', '.git', 'venv', '.venv'):
                        continue
                    child_path = os.path.join(path, entry)
                    child_node = self._build_tree(child_path, node)
                    node.children.append(child_node)
            except PermissionError:
                pass
        
        return node
    
    def _render_tree(self, node: FileNode, level: int):
        """渲染文件树"""
        if node is None:
            return
        
        # 创建行
        row = ctk.CTkFrame(self.tree_container, fg_color="transparent", height=26)
        row.pack(fill="x", pady=1)
        row.pack_propagate(False)
        
        self._item_widgets[node.path] = row
        
        # 缩进
        indent = level * 16
        
        # 展开/折叠图标（仅目录）
        if node.is_dir and node.children:
            is_expanded = node.path in self._expanded_paths
            toggle_text = "▼" if is_expanded else "▶"
            toggle_btn = ctk.CTkButton(
                row, text=toggle_text, width=16, height=20,
                fg_color="transparent", hover_color=COLORS['highlight'],
                text_color=COLORS['text_secondary'],
                font=ctk.CTkFont(size=8),
                command=lambda p=node.path: self._toggle_expand(p)
            )
            toggle_btn.pack(side="left", padx=(indent, 0))
        else:
            # 占位
            spacer = ctk.CTkFrame(row, width=indent + 16, height=1, fg_color="transparent")
            spacer.pack(side="left")
        
        # 图标
        icon = self._get_icon(node)
        icon_label = ctk.CTkLabel(
            row, text=icon, width=20,
            font=ctk.CTkFont(size=12)
        )
        icon_label.pack(side="left", padx=(2, 4))
        
        # 文件名
        name_label = ctk.CTkLabel(
            row, text=node.name,
            font=ctk.CTkFont(size=11),
            text_color=COLORS['text_primary'],
            anchor="w"
        )
        name_label.pack(side="left", fill="x", expand=True)
        
        # 绑定事件
        for widget in [row, icon_label, name_label]:
            widget.bind('<Button-1>', lambda e, n=node: self._on_click(n))
            widget.bind('<Double-Button-1>', lambda e, n=node: self._on_double_click(n))
            widget.bind('<Button-3>', lambda e, n=node: self._show_context_menu(e, n))
            widget.bind('<Enter>', lambda e, r=row: self._on_hover(r, True))
            widget.bind('<Leave>', lambda e, r=row: self._on_hover(r, False))
        
        # 渲染子节点（如果展开）
        if node.is_dir and node.path in self._expanded_paths:
            for child in node.children:
                self._render_tree(child, level + 1)
    
    def _get_icon(self, node: FileNode) -> str:
        """获取文件图标"""
        if node.is_dir:
            return self._icons['folder_open'] if node.path in self._expanded_paths else self._icons['folder']
        
        ext = os.path.splitext(node.name)[1].lower()
        ext_map = {
            '.md': 'md', '.markdown': 'md',
            '.txt': 'txt',
            '.py': 'py',
            '.json': 'json',
            '.html': 'html', '.htm': 'html',
            '.css': 'css',
            '.js': 'js',
            '.png': 'image', '.jpg': 'image', '.jpeg': 'image', '.gif': 'image', '.svg': 'image',
        }
        icon_key = ext_map.get(ext, 'default')
        return self._icons.get(icon_key, self._icons['default'])
    
    def _toggle_expand(self, path: str):
        """切换展开/折叠"""
        if path in self._expanded_paths:
            self._expanded_paths.discard(path)
        else:
            self._expanded_paths.add(path)
        self.refresh()
    
    def _on_click(self, node: FileNode):
        """单击事件"""
        # 更新选中状态
        if self.selected_path and self.selected_path in self._item_widgets:
            self._item_widgets[self.selected_path].configure(fg_color="transparent")
        
        self.selected_path = node.path
        if node.path in self._item_widgets:
            self._item_widgets[node.path].configure(fg_color=COLORS['highlight'])
    
    def _on_double_click(self, node: FileNode):
        """双击事件"""
        if node.is_dir:
            self._toggle_expand(node.path)
        else:
            # 打开文件
            if self.on_file_open:
                self.on_file_open(node.path)
    
    def _on_hover(self, row: ctk.CTkFrame, entering: bool):
        """悬停效果"""
        if row.cget('fg_color') != COLORS['highlight']:
            row.configure(fg_color=COLORS['bg_sidebar'] if entering else "transparent")
    
    def _show_context_menu(self, event, node: FileNode):
        """显示右键菜单"""
        menu = tk.Menu(self, tearoff=0)
        
        if node.is_dir:
            menu.add_command(label="📂 在资源管理器中打开", command=lambda: self._open_in_explorer(node.path))
            menu.add_command(label="🔄 刷新", command=self.refresh)
            menu.add_separator()
            menu.add_command(label="📄 新建文件", command=lambda: self._new_file(node.path))
            menu.add_command(label="📁 新建文件夹", command=lambda: self._new_folder(node.path))
        else:
            menu.add_command(label="📝 打开", command=lambda: self._open_file(node.path))
            menu.add_command(label="📂 在资源管理器中显示", command=lambda: self._open_in_explorer(os.path.dirname(node.path)))
            menu.add_separator()
            menu.add_command(label="📋 复制路径", command=lambda: self._copy_path(node.path))
            menu.add_command(label="📋 复制文件名", command=lambda: self._copy_path(node.name))
        
        menu.add_separator()
        menu.add_command(label="🗑️ 删除", command=lambda: self._delete(node))
        
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
    
    def _open_file(self, path: str):
        """打开文件"""
        if self.on_file_open:
            self.on_file_open(path)
    
    def _open_in_explorer(self, path: str):
        """在资源管理器中打开"""
        import subprocess
        import platform
        
        system = platform.system()
        try:
            if system == 'Windows':
                os.startfile(path)
            elif system == 'Darwin':
                subprocess.run(['open', path])
            else:
                subprocess.run(['xdg-open', path])
        except Exception as e:
            messagebox.showerror("错误", f"无法打开: {e}")
    
    def _copy_path(self, text: str):
        """复制到剪贴板"""
        self.clipboard_clear()
        self.clipboard_append(text)
    
    def _new_file(self, dir_path: str):
        """新建文件"""
        dialog = ctk.CTkInputDialog(
            text="输入文件名:",
            title="新建文件"
        )
        name = dialog.get_input()
        if name:
            file_path = os.path.join(dir_path, name)
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write('')
                self.refresh()
                if self.on_file_open:
                    self.on_file_open(file_path)
            except Exception as e:
                messagebox.showerror("错误", f"创建文件失败: {e}")
    
    def _new_folder(self, dir_path: str):
        """新建文件夹"""
        dialog = ctk.CTkInputDialog(
            text="输入文件夹名:",
            title="新建文件夹"
        )
        name = dialog.get_input()
        if name:
            folder_path = os.path.join(dir_path, name)
            try:
                os.makedirs(folder_path, exist_ok=True)
                self.refresh()
            except Exception as e:
                messagebox.showerror("错误", f"创建文件夹失败: {e}")
    
    def _delete(self, node: FileNode):
        """删除文件/文件夹"""
        msg = f"确定要删除 '{node.name}' 吗？"
        if node.is_dir:
            msg += "\n\n⚠️ 这将删除文件夹及其所有内容！"
        
        if messagebox.askyesno("确认删除", msg):
            try:
                if node.is_dir:
                    import shutil
                    shutil.rmtree(node.path)
                else:
                    os.remove(node.path)
                self.refresh()
            except Exception as e:
                messagebox.showerror("错误", f"删除失败: {e}")
    
    def collapse_all(self):
        """折叠全部"""
        self._expanded_paths.clear()
        if self.root_path:
            self._expanded_paths.add(self.root_path)
        self.refresh()
    
    def expand_to(self, path: str):
        """展开到指定路径"""
        if not self.root_path:
            return
        
        # 展开所有父目录
        current = path
        while current and current != self.root_path:
            parent = os.path.dirname(current)
            if parent:
                self._expanded_paths.add(parent)
            current = parent
        
        self.refresh()
        
        # 选中目标
        self.selected_path = path
        if path in self._item_widgets:
            self._item_widgets[path].configure(fg_color=COLORS['highlight'])


class FolderViewFeature:
    """文件夹视图功能"""
    
    def __init__(self, app):
        self.app = app
        self.folder_view: Optional[FolderTreeView] = None
    
    def create_view(self, parent) -> FolderTreeView:
        """创建文件夹视图"""
        self.folder_view = FolderTreeView(
            parent,
            on_file_open=self._on_file_open,
            fg_color="transparent"
        )
        return self.folder_view
    
    def _on_file_open(self, path: str):
        """打开文件回调"""
        # 只打开 Markdown 和文本文件
        ext = os.path.splitext(path)[1].lower()
        if ext in ('.md', '.markdown', '.txt', '.text'):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # 使用标签管理器打开
                if hasattr(self.app, 'tab_manager'):
                    self.app.tab_manager.open_file_in_tab(path, content)
                else:
                    # 直接设置内容
                    self.app.input_text.delete("1.0", "end")
                    self.app.input_text.insert("1.0", content)
                    self.app.current_file = path
                    self.app.on_text_change(None)
                
                self.app.update_status(f"已打开: {os.path.basename(path)}")
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件: {e}")
        else:
            # 其他文件用系统默认程序打开
            import subprocess
            import platform
            
            system = platform.system()
            try:
                if system == 'Windows':
                    os.startfile(path)
                elif system == 'Darwin':
                    subprocess.run(['open', path])
                else:
                    subprocess.run(['xdg-open', path])
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件: {e}")
    
    def open_folder(self, path: str = None):
        """打开文件夹"""
        if self.folder_view:
            self.folder_view.open_folder(path)
    
    def refresh(self):
        """刷新"""
        if self.folder_view:
            self.folder_view.refresh()
