# -*- coding: utf-8 -*-
"""协作版本历史模块"""

import hashlib
import time
from datetime import datetime
from typing import Optional, List, Dict
from dataclasses import dataclass, field

try:
    import customtkinter as ctk
except ImportError:
    ctk = None

from .theme import get_colors


@dataclass
class DocumentVersion:
    """文档版本"""
    version_id: str
    content: str
    author_id: str
    author_name: str
    timestamp: float
    description: str = ""
    content_hash: str = ""
    
    def __post_init__(self):
        if not self.content_hash:
            self.content_hash = hashlib.md5(self.content.encode()).hexdigest()[:8]
    
    @property
    def time_str(self) -> str:
        """格式化时间"""
        dt = datetime.fromtimestamp(self.timestamp)
        now = datetime.now()
        
        if dt.date() == now.date():
            return f"今天 {dt.strftime('%H:%M')}"
        elif (now - dt).days == 1:
            return f"昨天 {dt.strftime('%H:%M')}"
        else:
            return dt.strftime("%m-%d %H:%M")
    
    @property
    def size_str(self) -> str:
        """内容大小"""
        size = len(self.content.encode('utf-8'))
        if size < 1024:
            return f"{size} B"
        elif size < 1024 * 1024:
            return f"{size / 1024:.1f} KB"
        else:
            return f"{size / (1024 * 1024):.1f} MB"
    
    def to_dict(self) -> dict:
        return {
            'version_id': self.version_id,
            'content': self.content,
            'author_id': self.author_id,
            'author_name': self.author_name,
            'timestamp': self.timestamp,
            'description': self.description,
            'content_hash': self.content_hash,
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'DocumentVersion':
        return cls(
            version_id=data.get('version_id', ''),
            content=data.get('content', ''),
            author_id=data.get('author_id', ''),
            author_name=data.get('author_name', ''),
            timestamp=data.get('timestamp', time.time()),
            description=data.get('description', ''),
            content_hash=data.get('content_hash', ''),
        )


class VersionHistoryManager:
    """版本历史管理器"""
    
    def __init__(self, max_versions: int = 50):
        self._versions: List[DocumentVersion] = []
        self._max_versions = max_versions
        self._version_counter = 0
        self._last_content_hash = ""
    
    def add_version(self, content: str, author_id: str, author_name: str,
                    description: str = "") -> Optional[DocumentVersion]:
        """添加新版本"""
        # 计算内容哈希，避免重复保存相同内容
        content_hash = hashlib.md5(content.encode()).hexdigest()[:8]
        if content_hash == self._last_content_hash:
            return None
        
        self._version_counter += 1
        version = DocumentVersion(
            version_id=f"v{self._version_counter}",
            content=content,
            author_id=author_id,
            author_name=author_name,
            timestamp=time.time(),
            description=description or f"版本 {self._version_counter}",
            content_hash=content_hash,
        )
        
        self._versions.append(version)
        self._last_content_hash = content_hash
        
        # 限制版本数量
        if len(self._versions) > self._max_versions:
            self._versions = self._versions[-self._max_versions:]
        
        return version
    
    def get_versions(self, limit: int = None) -> List[DocumentVersion]:
        """获取版本列表"""
        versions = list(reversed(self._versions))
        if limit:
            return versions[:limit]
        return versions
    
    def get_version(self, version_id: str) -> Optional[DocumentVersion]:
        """获取指定版本"""
        for v in self._versions:
            if v.version_id == version_id:
                return v
        return None
    
    def clear(self):
        """清空历史"""
        self._versions = []
        self._version_counter = 0
        self._last_content_hash = ""


class VersionHistoryPanel:
    """版本历史面板"""
    
    def __init__(self, app, feature):
        self.app = app
        self.feature = feature
        self.dialog: Optional[ctk.CTkToplevel] = None
        self._on_restore_callback = None
    
    def set_on_restore(self, callback):
        """设置恢复回调"""
        self._on_restore_callback = callback
    
    def show(self):
        """显示版本历史对话框"""
        if self.dialog and self.dialog.winfo_exists():
            self.dialog.lift()
            return
        
        colors = get_colors()
        
        self.dialog = ctk.CTkToplevel(self.app)
        self.dialog.title("📜 版本历史")
        self.dialog.geometry("500x600")
        self.dialog.transient(self.app)
        self.dialog.configure(fg_color=colors['background'])
        
        # 标题
        header = ctk.CTkFrame(self.dialog, fg_color=colors['background_secondary'], height=55)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        header_left = ctk.CTkFrame(header, fg_color="transparent")
        header_left.pack(side="left", padx=20, pady=12)
        
        ctk.CTkLabel(header_left, text="📜 版本历史",
                    font=ctk.CTkFont(size=16, weight="bold"),
                    text_color=colors['text']).pack(side="left")
        
        # 版本数量
        versions = self.feature.version_manager.get_versions() if hasattr(self.feature, 'version_manager') else []
        if versions:
            ctk.CTkLabel(header_left, text=f" · {len(versions)}个版本",
                        font=ctk.CTkFont(size=11),
                        text_color=colors['text_muted']).pack(side="left")
        
        # 版本列表
        scroll = ctk.CTkScrollableFrame(self.dialog, fg_color=colors['background'])
        scroll.pack(fill="both", expand=True, padx=15, pady=15)
        
        if not versions:
            # 空状态
            empty_frame = ctk.CTkFrame(scroll, fg_color="transparent")
            empty_frame.pack(expand=True, pady=80)
            
            ctk.CTkLabel(empty_frame, text="📭",
                        font=ctk.CTkFont(size=48)).pack()
            ctk.CTkLabel(empty_frame, text="暂无版本历史",
                        font=ctk.CTkFont(size=14),
                        text_color=colors['text_secondary']).pack(pady=10)
            ctk.CTkLabel(empty_frame, text="文档修改后会自动保存版本",
                        font=ctk.CTkFont(size=11),
                        text_color=colors['text_muted']).pack()
            return
        
        for version in versions:
            self._create_version_card(scroll, version, colors)
    
    def _create_version_card(self, parent, version: DocumentVersion, colors: dict):
        """创建版本卡片"""
        card = ctk.CTkFrame(parent, fg_color=colors['surface'], corner_radius=10)
        card.pack(fill="x", pady=5, padx=5)
        
        # 左侧信息
        info = ctk.CTkFrame(card, fg_color="transparent")
        info.pack(side="left", fill="x", expand=True, padx=15, pady=12)
        
        # 版本号和时间
        header = ctk.CTkFrame(info, fg_color="transparent")
        header.pack(anchor="w")
        
        ctk.CTkLabel(header, text=version.version_id,
                    font=ctk.CTkFont(size=12, weight="bold"),
                    text_color=colors['primary']).pack(side="left")
        
        ctk.CTkLabel(header, text=f" · {version.time_str}",
                    font=ctk.CTkFont(size=11),
                    text_color=colors['text_secondary']).pack(side="left")
        
        # 作者和大小
        meta = ctk.CTkFrame(info, fg_color="transparent")
        meta.pack(anchor="w", pady=(3, 0))
        
        ctk.CTkLabel(meta, text=f"👤 {version.author_name}",
                    font=ctk.CTkFont(size=10),
                    text_color=colors['text_muted']).pack(side="left")
        
        ctk.CTkLabel(meta, text=f" · {version.size_str}",
                    font=ctk.CTkFont(size=10),
                    text_color=colors['text_muted']).pack(side="left")
        
        ctk.CTkLabel(meta, text=f" · #{version.content_hash}",
                    font=ctk.CTkFont(size=10),
                    text_color=colors['text_muted']).pack(side="left")
        
        # 描述
        if version.description:
            ctk.CTkLabel(info, text=version.description,
                        font=ctk.CTkFont(size=10),
                        text_color=colors['text_secondary']).pack(anchor="w", pady=(3, 0))
        
        # 右侧按钮
        btn_frame = ctk.CTkFrame(card, fg_color="transparent")
        btn_frame.pack(side="right", padx=10)
        
        # 预览按钮
        preview_btn = ctk.CTkButton(
            btn_frame, text="👁", width=32, height=32,
            fg_color=colors['gray_light'], hover_color=colors['surface_hover'],
            corner_radius=16,
            command=lambda v=version: self._preview_version(v)
        )
        preview_btn.pack(side="left", padx=3)
        
        # 恢复按钮
        restore_btn = ctk.CTkButton(
            btn_frame, text="恢复", width=60, height=32,
            fg_color=colors['primary'], hover_color=colors['primary_dark'],
            corner_radius=16, font=ctk.CTkFont(size=11),
            command=lambda v=version: self._restore_version(v)
        )
        restore_btn.pack(side="left", padx=3)
    
    def _preview_version(self, version: DocumentVersion):
        """预览版本"""
        colors = get_colors()
        
        preview = ctk.CTkToplevel(self.app)
        preview.title(f"预览 - {version.version_id}")
        preview.geometry("600x500")
        preview.transient(self.dialog)
        
        # 信息栏
        info = ctk.CTkFrame(preview, fg_color=colors['background_secondary'], height=40)
        info.pack(fill="x")
        info.pack_propagate(False)
        
        ctk.CTkLabel(info, text=f"{version.version_id} · {version.author_name} · {version.time_str}",
                    font=ctk.CTkFont(size=11),
                    text_color=colors['text_secondary']).pack(side="left", padx=15, pady=10)
        
        # 内容
        text = ctk.CTkTextbox(preview, fg_color=colors['surface'],
                             text_color=colors['text'], wrap="word")
        text.pack(fill="both", expand=True, padx=10, pady=10)
        text.insert("1.0", version.content)
        text.configure(state="disabled")
    
    def _restore_version(self, version: DocumentVersion):
        """恢复版本"""
        from tkinter import messagebox
        
        if messagebox.askyesno("确认恢复", f"确定要恢复到 {version.version_id} 吗？\n当前内容将被替换。"):
            if self._on_restore_callback:
                self._on_restore_callback(version.content)
            
            if self.dialog:
                self.dialog.destroy()
