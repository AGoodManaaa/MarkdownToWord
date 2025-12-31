# -*- coding: utf-8 -*-
"""数据库功能面板模块"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Optional, List
from pathlib import Path

try:
    import customtkinter as ctk
except ImportError:
    ctk = None

from .vault import VaultManager, VaultInfo, DocumentInfo, VaultError
from .search import SearchEngine, SearchResult
from .tags import TagManager, TagInfo
from .links import LinkManager
from .graph import GraphView, GraphData


class DatabaseFeature:
    """数据库功能入口"""
    
    def __init__(self, app):
        self.app = app
        self.vault_manager: Optional[VaultManager] = None
        self.search_engine: Optional[SearchEngine] = None
        self.tag_manager: Optional[TagManager] = None
        self.link_manager: Optional[LinkManager] = None
        self.graph_view: Optional[GraphView] = None
        
        self._search_panel: Optional['SearchPanel'] = None
        self._tag_panel: Optional['TagPanel'] = None
        self._backlink_panel: Optional['BacklinkPanel'] = None
        self._graph_dialog: Optional['GraphDialog'] = None
    
    def initialize(self, vault_path: str) -> VaultInfo:
        """初始化数据库功能
        
        Args:
            vault_path: 文档库路径
            
        Returns:
            VaultInfo 文档库信息
        """
        # 创建数据库文件路径
        db_path = os.path.join(vault_path, '.vault.db')
        
        # 初始化各组件
        self.vault_manager = VaultManager(db_path)
        vault_info = self.vault_manager.open_vault(vault_path)
        
        self.search_engine = SearchEngine(self.vault_manager)
        self.tag_manager = TagManager(self.vault_manager)
        self.link_manager = LinkManager(self.vault_manager)
        self.graph_view = GraphView(self.link_manager)
        
        # 启动文件监控
        self.vault_manager.watch_changes(self._on_file_change)
        
        return vault_info
    
    def _on_file_change(self, path: str, event_type: str) -> None:
        """文件变化回调"""
        # 刷新断开链接状态
        if self.link_manager:
            self.link_manager.refresh_broken_links()
        
        # 更新 UI（如果打开）
        if self._backlink_panel:
            self._backlink_panel.refresh()
    
    def show_vault_selector(self) -> None:
        """显示文档库选择对话框"""
        if ctk is None:
            return
        
        dialog = VaultSelectorDialog(self.app, self)
        dialog.show()
    
    def show_search_panel(self) -> None:
        """显示搜索面板"""
        if not self.vault_manager:
            messagebox.showinfo("提示", "请先打开文档库")
            self.show_vault_selector()
            return
        
        if self._search_panel is None:
            self._search_panel = SearchPanel(self.app, self)
        self._search_panel.show()
    
    def show_tag_panel(self) -> None:
        """显示标签面板"""
        if not self.vault_manager:
            messagebox.showinfo("提示", "请先打开文档库")
            self.show_vault_selector()
            return
        
        if self._tag_panel is None:
            self._tag_panel = TagPanel(self.app, self)
        self._tag_panel.show()
    
    def show_backlink_panel(self, document_path: str) -> None:
        """显示反向链接面板"""
        if not self.vault_manager:
            return
        
        if self._backlink_panel is None:
            self._backlink_panel = BacklinkPanel(self.app, self)
        self._backlink_panel.show(document_path)
    
    def show_graph_view(self) -> None:
        """显示知识图谱"""
        if not self.vault_manager:
            messagebox.showinfo("提示", "请先打开文档库")
            self.show_vault_selector()
            return
        
        if self._graph_dialog is None:
            self._graph_dialog = GraphDialog(self.app, self)
        self._graph_dialog.show()
    
    def search(self, query: str) -> List[SearchResult]:
        """执行搜索"""
        if self.search_engine:
            return self.search_engine.search(query)
        return []
    
    def get_link_suggestions(self, prefix: str) -> List[str]:
        """获取链接建议"""
        if self.link_manager:
            return self.link_manager.get_link_suggestions(prefix)
        return []
    
    def close(self) -> None:
        """关闭数据库功能"""
        if self.vault_manager:
            self.vault_manager.close()
            self.vault_manager = None


class VaultSelectorDialog:
    """文档库选择对话框"""
    
    def __init__(self, app, feature: DatabaseFeature):
        self.app = app
        self.feature = feature
        self.dialog: Optional[ctk.CTkToplevel] = None
    
    def show(self) -> None:
        """显示对话框"""
        if self.dialog is not None and self.dialog.winfo_exists():
            self.dialog.focus()
            return
        
        self.dialog = ctk.CTkToplevel(self.app)
        self.dialog.title("📚 选择文档库")
        self.dialog.geometry("400x200")
        self.dialog.resizable(False, False)
        
        # 说明
        label = ctk.CTkLabel(
            self.dialog,
            text="选择一个文件夹作为文档库\n系统将索引其中的所有 Markdown 文件",
            font=ctk.CTkFont(size=12)
        )
        label.pack(pady=20)
        
        # 按钮
        btn_frame = ctk.CTkFrame(self.dialog, fg_color="transparent")
        btn_frame.pack(pady=20)
        
        open_btn = ctk.CTkButton(
            btn_frame,
            text="📂 打开文档库",
            command=self._on_open,
            width=150
        )
        open_btn.pack(side="left", padx=10)
        
        create_btn = ctk.CTkButton(
            btn_frame,
            text="➕ 创建新文档库",
            command=self._on_create,
            width=150
        )
        create_btn.pack(side="left", padx=10)
    
    def _on_open(self) -> None:
        """打开文档库"""
        path = filedialog.askdirectory(title="选择文档库文件夹")
        if path:
            try:
                vault_info = self.feature.initialize(path)
                messagebox.showinfo(
                    "成功",
                    f"已打开文档库: {vault_info.name}\n"
                    f"文档数量: {vault_info.file_count}"
                )
                self.dialog.destroy()
            except VaultError as e:
                messagebox.showerror("错误", str(e))
    
    def _on_create(self) -> None:
        """创建新文档库"""
        path = filedialog.askdirectory(title="选择新文档库位置")
        if path:
            try:
                vault_info = self.feature.initialize(path)
                messagebox.showinfo(
                    "成功",
                    f"已创建文档库: {vault_info.name}"
                )
                self.dialog.destroy()
            except VaultError as e:
                messagebox.showerror("错误", str(e))


class SearchPanel:
    """搜索面板"""
    
    def __init__(self, app, feature: DatabaseFeature):
        self.app = app
        self.feature = feature
        self.dialog: Optional[ctk.CTkToplevel] = None
    
    def show(self) -> None:
        """显示搜索面板"""
        if self.dialog is not None and self.dialog.winfo_exists():
            self.dialog.focus()
            return
        
        self.dialog = ctk.CTkToplevel(self.app)
        self.dialog.title("🔍 全局搜索")
        self.dialog.geometry("600x500")
        
        # 搜索框
        search_frame = ctk.CTkFrame(self.dialog, fg_color="transparent")
        search_frame.pack(fill="x", padx=15, pady=15)
        
        self.search_entry = ctk.CTkEntry(
            search_frame,
            placeholder_text="输入搜索关键词...",
            width=500
        )
        self.search_entry.pack(side="left", fill="x", expand=True)
        self.search_entry.bind('<Return>', lambda e: self._do_search())
        self.search_entry.bind('<KeyRelease>', self._on_key_release)
        
        search_btn = ctk.CTkButton(
            search_frame,
            text="搜索",
            command=self._do_search,
            width=80
        )
        search_btn.pack(side="left", padx=(10, 0))
        
        # 结果列表
        self.result_frame = ctk.CTkScrollableFrame(self.dialog)
        self.result_frame.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        # 状态栏
        self.status_label = ctk.CTkLabel(
            self.dialog,
            text="输入关键词开始搜索",
            font=ctk.CTkFont(size=11)
        )
        self.status_label.pack(pady=(0, 10))
    
    def _on_key_release(self, event) -> None:
        """按键释放事件（实时搜索）"""
        query = self.search_entry.get()
        if len(query) >= 2:
            self._do_search()
    
    def _do_search(self) -> None:
        """执行搜索"""
        query = self.search_entry.get().strip()
        if not query:
            return
        
        # 清空结果
        for widget in self.result_frame.winfo_children():
            widget.destroy()
        
        # 执行搜索
        results = self.feature.search(query)
        
        # 显示结果
        for result in results:
            self._create_result_item(result)
        
        self.status_label.configure(text=f"找到 {len(results)} 个结果")
    
    def _create_result_item(self, result: SearchResult) -> None:
        """创建结果项"""
        item_frame = ctk.CTkFrame(self.result_frame)
        item_frame.pack(fill="x", pady=5)
        
        # 标题
        title_label = ctk.CTkLabel(
            item_frame,
            text=result.document.title,
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w"
        )
        title_label.pack(fill="x", padx=10, pady=(5, 0))
        
        # 路径
        path_label = ctk.CTkLabel(
            item_frame,
            text=str(result.document.path),
            font=ctk.CTkFont(size=10),
            text_color="gray",
            anchor="w"
        )
        path_label.pack(fill="x", padx=10)
        
        # 匹配片段
        if result.snippet:
            snippet_label = ctk.CTkLabel(
                item_frame,
                text=result.snippet[:100] + "..." if len(result.snippet) > 100 else result.snippet,
                font=ctk.CTkFont(size=11),
                anchor="w",
                wraplength=550
            )
            snippet_label.pack(fill="x", padx=10, pady=(0, 5))
        
        # 点击打开
        item_frame.bind('<Button-1>', lambda e, p=result.document.path: self._open_document(p))
        for child in item_frame.winfo_children():
            child.bind('<Button-1>', lambda e, p=result.document.path: self._open_document(p))
    
    def _open_document(self, path) -> None:
        """打开文档"""
        if self.feature.vault_manager and self.feature.vault_manager.current_vault:
            full_path = self.feature.vault_manager.current_vault / path
            if full_path.exists():
                # 在应用中打开文件
                if hasattr(self.app, 'file_ops'):
                    self.app.file_ops.open_file(str(full_path))


class TagPanel:
    """标签面板"""
    
    def __init__(self, app, feature: DatabaseFeature):
        self.app = app
        self.feature = feature
        self.dialog: Optional[ctk.CTkToplevel] = None
    
    def show(self) -> None:
        """显示标签面板"""
        if self.dialog is not None and self.dialog.winfo_exists():
            self.dialog.focus()
            return
        
        self.dialog = ctk.CTkToplevel(self.app)
        self.dialog.title("🏷️ 标签管理")
        self.dialog.geometry("400x500")
        
        # 标签列表
        self.tag_frame = ctk.CTkScrollableFrame(self.dialog)
        self.tag_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        self._refresh_tags()
    
    def _refresh_tags(self) -> None:
        """刷新标签列表"""
        # 清空
        for widget in self.tag_frame.winfo_children():
            widget.destroy()
        
        if not self.feature.tag_manager:
            return
        
        tags = self.feature.tag_manager.get_all_tags()
        
        for tag in tags:
            self._create_tag_item(tag)
    
    def _create_tag_item(self, tag: TagInfo) -> None:
        """创建标签项"""
        item_frame = ctk.CTkFrame(self.tag_frame)
        item_frame.pack(fill="x", pady=2)
        
        # 标签名
        name_label = ctk.CTkLabel(
            item_frame,
            text=f"#{tag.name}",
            font=ctk.CTkFont(size=12),
            anchor="w"
        )
        name_label.pack(side="left", padx=10, pady=5)
        
        # 文档数量
        count_label = ctk.CTkLabel(
            item_frame,
            text=f"({tag.count})",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        )
        count_label.pack(side="left")
        
        # 点击过滤
        item_frame.bind('<Button-1>', lambda e, t=tag.name: self._filter_by_tag(t))
        name_label.bind('<Button-1>', lambda e, t=tag.name: self._filter_by_tag(t))
    
    def _filter_by_tag(self, tag: str) -> None:
        """按标签过滤"""
        if self.feature.tag_manager:
            docs = self.feature.tag_manager.get_documents_by_tag(tag)
            messagebox.showinfo("标签文档", f"#{tag} 包含 {len(docs)} 个文档")


class BacklinkPanel:
    """反向链接面板"""
    
    def __init__(self, app, feature: DatabaseFeature):
        self.app = app
        self.feature = feature
        self.dialog: Optional[ctk.CTkToplevel] = None
        self.current_path: Optional[str] = None
    
    def show(self, document_path: str) -> None:
        """显示反向链接面板"""
        self.current_path = document_path
        
        if self.dialog is not None and self.dialog.winfo_exists():
            self.refresh()
            self.dialog.focus()
            return
        
        self.dialog = ctk.CTkToplevel(self.app)
        self.dialog.title("🔗 反向链接")
        self.dialog.geometry("350x400")
        
        # 当前文档
        self.doc_label = ctk.CTkLabel(
            self.dialog,
            text=f"📄 {Path(document_path).stem}",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.doc_label.pack(pady=10)
        
        # 链接列表
        self.link_frame = ctk.CTkScrollableFrame(self.dialog)
        self.link_frame.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        self.refresh()
    
    def refresh(self) -> None:
        """刷新链接列表"""
        if not self.current_path or not self.feature.link_manager:
            return
        
        # 清空
        for widget in self.link_frame.winfo_children():
            widget.destroy()
        
        # 获取反向链接
        backlinks = self.feature.link_manager.get_backlinks(self.current_path)
        
        if not backlinks:
            label = ctk.CTkLabel(
                self.link_frame,
                text="没有反向链接",
                text_color="gray"
            )
            label.pack(pady=20)
            return
        
        for link in backlinks:
            self._create_link_item(link)
    
    def _create_link_item(self, link) -> None:
        """创建链接项"""
        item_frame = ctk.CTkFrame(self.link_frame)
        item_frame.pack(fill="x", pady=2)
        
        # 源文档名
        name_label = ctk.CTkLabel(
            item_frame,
            text=f"📄 {Path(link.source).stem}",
            font=ctk.CTkFont(size=12),
            anchor="w"
        )
        name_label.pack(fill="x", padx=10, pady=5)
        
        # 点击打开
        item_frame.bind('<Button-1>', lambda e, p=link.source: self._open_document(p))
        name_label.bind('<Button-1>', lambda e, p=link.source: self._open_document(p))
    
    def _open_document(self, path) -> None:
        """打开文档"""
        if self.feature.vault_manager and self.feature.vault_manager.current_vault:
            full_path = self.feature.vault_manager.current_vault / path
            if full_path.exists():
                if hasattr(self.app, 'file_ops'):
                    self.app.file_ops.open_file(str(full_path))


class GraphDialog:
    """知识图谱对话框"""
    
    def __init__(self, app, feature: DatabaseFeature):
        self.app = app
        self.feature = feature
        self.dialog: Optional[ctk.CTkToplevel] = None
        self.graph_data: Optional[GraphData] = None
    
    def show(self) -> None:
        """显示图谱对话框"""
        if self.dialog is not None and self.dialog.winfo_exists():
            self.dialog.focus()
            return
        
        self.dialog = ctk.CTkToplevel(self.app)
        self.dialog.title("🕸️ 知识图谱")
        self.dialog.geometry("900x700")
        
        # 工具栏
        toolbar = ctk.CTkFrame(self.dialog, fg_color="transparent", height=40)
        toolbar.pack(fill="x", padx=15, pady=10)
        
        refresh_btn = ctk.CTkButton(
            toolbar,
            text="🔄 刷新",
            command=self._refresh_graph,
            width=80
        )
        refresh_btn.pack(side="left")
        
        # 统计信息
        self.stats_label = ctk.CTkLabel(
            toolbar,
            text="",
            font=ctk.CTkFont(size=11)
        )
        self.stats_label.pack(side="right")
        
        # 画布
        self.canvas = tk.Canvas(
            self.dialog,
            bg="#1a1a2e",
            highlightthickness=0
        )
        self.canvas.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        # 绑定事件
        self.canvas.bind('<Configure>', self._on_resize)
        self.canvas.bind('<Button-1>', self._on_click)
        self.canvas.bind('<Double-Button-1>', self._on_double_click)
        
        # 初始加载
        self._refresh_graph()
    
    def _refresh_graph(self) -> None:
        """刷新图谱"""
        if not self.feature.vault_manager or not self.feature.graph_view:
            return
        
        # 获取所有文档
        documents = self.feature.vault_manager.get_all_documents()
        
        # 构建图
        width = self.canvas.winfo_width() or 800
        height = self.canvas.winfo_height() or 600
        
        self.graph_data = self.feature.graph_view.build_graph(documents)
        self.graph_data = self.feature.graph_view.calculate_layout(
            self.graph_data, width, height
        )
        
        # 绘制
        self._draw_graph()
        
        # 更新统计
        stats = self.feature.graph_view.get_statistics(self.graph_data)
        self.stats_label.configure(
            text=f"节点: {stats['node_count']} | 边: {stats['edge_count']} | 平均度: {stats['avg_degree']}"
        )
    
    def _draw_graph(self) -> None:
        """绘制图谱"""
        self.canvas.delete("all")
        
        if not self.graph_data:
            return
        
        # 绘制边
        for edge in self.graph_data.edges:
            source = self.graph_data.get_node(edge.source)
            target = self.graph_data.get_node(edge.target)
            
            if source and target:
                self.canvas.create_line(
                    source.x, source.y,
                    target.x, target.y,
                    fill="#4a4a6a",
                    width=1
                )
        
        # 绘制节点
        for node in self.graph_data.nodes:
            # 节点圆
            r = node.size / 2
            self.canvas.create_oval(
                node.x - r, node.y - r,
                node.x + r, node.y + r,
                fill=node.color,
                outline="white",
                width=1,
                tags=f"node_{node.id}"
            )
            
            # 标签
            self.canvas.create_text(
                node.x, node.y + r + 10,
                text=node.label[:15] + "..." if len(node.label) > 15 else node.label,
                fill="white",
                font=("Microsoft YaHei", 9),
                tags=f"label_{node.id}"
            )
    
    def _on_resize(self, event) -> None:
        """窗口大小改变"""
        if self.graph_data and self.feature.graph_view:
            self.graph_data = self.feature.graph_view.calculate_layout(
                self.graph_data, event.width, event.height
            )
            self._draw_graph()
    
    def _on_click(self, event) -> None:
        """单击事件"""
        # 查找点击的节点
        items = self.canvas.find_overlapping(
            event.x - 5, event.y - 5,
            event.x + 5, event.y + 5
        )
        
        for item in items:
            tags = self.canvas.gettags(item)
            for tag in tags:
                if tag.startswith("node_"):
                    node_id = tag[5:]
                    self._highlight_node(node_id)
                    return
    
    def _on_double_click(self, event) -> None:
        """双击事件"""
        items = self.canvas.find_overlapping(
            event.x - 5, event.y - 5,
            event.x + 5, event.y + 5
        )
        
        for item in items:
            tags = self.canvas.gettags(item)
            for tag in tags:
                if tag.startswith("node_"):
                    node_id = tag[5:]
                    self._open_document(node_id)
                    return
    
    def _highlight_node(self, node_id: str) -> None:
        """高亮节点及其邻居"""
        if not self.graph_data or not self.feature.graph_view:
            return
        
        neighbors = self.feature.graph_view.get_neighbors(self.graph_data, node_id)
        neighbors.add(node_id)
        
        # 重绘，高亮相关节点
        self._draw_graph()
        
        for node in self.graph_data.nodes:
            if node.id in neighbors:
                r = node.size / 2 + 3
                self.canvas.create_oval(
                    node.x - r, node.y - r,
                    node.x + r, node.y + r,
                    outline="#FFD700",
                    width=2
                )
    
    def _open_document(self, path: str) -> None:
        """打开文档"""
        if self.feature.vault_manager and self.feature.vault_manager.current_vault:
            full_path = self.feature.vault_manager.current_vault / path
            if full_path.exists():
                if hasattr(self.app, 'file_ops'):
                    self.app.file_ops.open_file(str(full_path))
