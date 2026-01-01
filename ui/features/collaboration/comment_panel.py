# -*- coding: utf-8 -*-
"""评论侧边栏面板模块 - 腾讯文档风格

Features:
- 右侧滑出评论面板
- 评论气泡组件
- 表情回复功能
- 评论筛选
- @提及自动补全
"""

from dataclasses import dataclass, field
from typing import List, Optional, Dict, Callable, Any
from datetime import datetime
import tkinter as tk

try:
    import customtkinter as ctk
    CTK_AVAILABLE = True
except ImportError:
    ctk = None
    CTK_AVAILABLE = False


# 表情回复列表
REACTIONS = [
    {"emoji": "👍", "name": "赞"},
    {"emoji": "❤️", "name": "爱心"},
    {"emoji": "😄", "name": "开心"},
    {"emoji": "🎉", "name": "庆祝"},
    {"emoji": "🤔", "name": "思考"},
    {"emoji": "👀", "name": "关注"},
    {"emoji": "🔥", "name": "火热"},
    {"emoji": "✅", "name": "完成"},
]

# 配色
COLORS = {
    'primary': '#10b981',
    'primary_light': '#d1fae5',
    'secondary': '#3b82f6',
    'secondary_light': '#dbeafe',
    'danger': '#ef4444',
    'warning': '#f59e0b',
    'surface': '#ffffff',
    'surface_hover': '#f3f4f6',
    'border': '#e5e7eb',
    'text': '#1f2937',
    'text_secondary': '#6b7280',
    'text_muted': '#9ca3af',
    'background': '#f9fafb',
}


@dataclass
class CommentReaction:
    """评论表情回复"""
    emoji: str
    user_id: str
    user_name: str
    timestamp: datetime = field(default_factory=datetime.now)


@dataclass
class CommentData:
    """评论数据"""
    id: str
    thread_id: str
    author_id: str
    author_name: str
    author_color: str
    content: str
    created_at: datetime
    resolved: bool = False
    reactions: List[CommentReaction] = field(default_factory=list)
    replies: List['CommentData'] = field(default_factory=list)
    
    @property
    def reaction_summary(self) -> Dict[str, int]:
        """统计表情回复"""
        summary = {}
        for r in self.reactions:
            summary[r.emoji] = summary.get(r.emoji, 0) + 1
        return summary


class CommentBubble(ctk.CTkFrame if CTK_AVAILABLE else object):
    """评论气泡组件 - 腾讯文档风格"""
    
    def __init__(self, parent, comment: CommentData, 
                 on_reply: Callable = None,
                 on_resolve: Callable = None,
                 on_reaction: Callable = None,
                 current_user_id: str = None,
                 **kwargs):
        if not CTK_AVAILABLE:
            return
            
        super().__init__(parent, fg_color=COLORS['surface'], corner_radius=12,
                        border_width=1, border_color=COLORS['border'], **kwargs)
        
        self.comment = comment
        self.on_reply = on_reply
        self.on_resolve = on_resolve
        self.on_reaction = on_reaction
        self.current_user_id = current_user_id
        self._reaction_frame = None
        self._reply_entry = None
        
        self._create_ui()
    
    def _create_ui(self):
        """创建评论气泡 UI"""
        # 头部：头像 + 用户名 + 时间
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=12, pady=(12, 6))
        
        # 头像
        avatar = ctk.CTkLabel(header, text=self.comment.author_name[0].upper(),
                             width=32, height=32, 
                             fg_color=self.comment.author_color or COLORS['primary'],
                             corner_radius=16, 
                             font=ctk.CTkFont(size=12, weight="bold"),
                             text_color="white")
        avatar.pack(side="left")
        
        # 用户信息
        info_frame = ctk.CTkFrame(header, fg_color="transparent")
        info_frame.pack(side="left", padx=10, fill="x", expand=True)
        
        ctk.CTkLabel(info_frame, text=self.comment.author_name,
                    font=ctk.CTkFont(size=12, weight="bold"),
                    text_color=COLORS['text']).pack(anchor="w")
        
        time_str = self.comment.created_at.strftime("%m-%d %H:%M")
        ctk.CTkLabel(info_frame, text=time_str,
                    font=ctk.CTkFont(size=10),
                    text_color=COLORS['text_muted']).pack(anchor="w")
        
        # 解决按钮（如果是主评论）
        if not self.comment.resolved:
            resolve_btn = ctk.CTkButton(header, text="✓", width=28, height=28,
                                       corner_radius=14,
                                       fg_color=COLORS['surface_hover'],
                                       hover_color=COLORS['primary_light'],
                                       text_color=COLORS['text_secondary'],
                                       font=ctk.CTkFont(size=12),
                                       command=self._on_resolve_click)
            resolve_btn.pack(side="right")
        else:
            ctk.CTkLabel(header, text="✅ 已解决",
                        font=ctk.CTkFont(size=10),
                        text_color=COLORS['primary']).pack(side="right")
        
        # 评论内容
        content_frame = ctk.CTkFrame(self, fg_color="transparent")
        content_frame.pack(fill="x", padx=12, pady=6)
        
        content_label = ctk.CTkLabel(content_frame, text=self.comment.content,
                                    font=ctk.CTkFont(size=12),
                                    text_color=COLORS['text'],
                                    wraplength=280, justify="left")
        content_label.pack(anchor="w")
        
        # 表情回复区域
        if self.comment.reactions:
            self._create_reaction_display()
        
        # 操作栏
        action_frame = ctk.CTkFrame(self, fg_color="transparent")
        action_frame.pack(fill="x", padx=12, pady=(6, 12))
        
        # 添加表情按钮
        emoji_btn = ctk.CTkButton(action_frame, text="😊", width=32, height=28,
                                 corner_radius=14,
                                 fg_color="transparent",
                                 hover_color=COLORS['surface_hover'],
                                 text_color=COLORS['text_secondary'],
                                 font=ctk.CTkFont(size=14),
                                 command=self._show_reaction_picker)
        emoji_btn.pack(side="left")
        
        # 回复按钮
        reply_btn = ctk.CTkButton(action_frame, text="💬 回复", width=60, height=28,
                                 corner_radius=14,
                                 fg_color="transparent",
                                 hover_color=COLORS['surface_hover'],
                                 text_color=COLORS['text_secondary'],
                                 font=ctk.CTkFont(size=11),
                                 command=self._show_reply_input)
        reply_btn.pack(side="left", padx=5)
        
        # 显示回复
        if self.comment.replies:
            self._create_replies_section()
    
    def _create_reaction_display(self):
        """显示表情回复统计"""
        self._reaction_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._reaction_frame.pack(fill="x", padx=12, pady=(0, 6))
        
        for emoji, count in self.comment.reaction_summary.items():
            pill = ctk.CTkFrame(self._reaction_frame, 
                               fg_color=COLORS['surface_hover'],
                               corner_radius=10)
            pill.pack(side="left", padx=2)
            
            ctk.CTkLabel(pill, text=f"{emoji} {count}",
                        font=ctk.CTkFont(size=11),
                        text_color=COLORS['text']).pack(padx=8, pady=2)
    
    def _create_replies_section(self):
        """显示回复列表"""
        replies_frame = ctk.CTkFrame(self, fg_color=COLORS['background'],
                                    corner_radius=8)
        replies_frame.pack(fill="x", padx=12, pady=(0, 10))
        
        for reply in self.comment.replies[:3]:  # 最多显示 3 条
            reply_item = ctk.CTkFrame(replies_frame, fg_color="transparent")
            reply_item.pack(fill="x", padx=10, pady=5)
            
            # 回复者头像（小尺寸）
            avatar = ctk.CTkLabel(reply_item, text=reply.author_name[0].upper(),
                                 width=24, height=24,
                                 fg_color=reply.author_color or COLORS['secondary'],
                                 corner_radius=12,
                                 font=ctk.CTkFont(size=9, weight="bold"),
                                 text_color="white")
            avatar.pack(side="left")
            
            # 回复内容
            ctk.CTkLabel(reply_item, 
                        text=f"{reply.author_name}: {reply.content}",
                        font=ctk.CTkFont(size=11),
                        text_color=COLORS['text'],
                        wraplength=240, justify="left").pack(side="left", padx=8)
        
        if len(self.comment.replies) > 3:
            ctk.CTkLabel(replies_frame, 
                        text=f"查看全部 {len(self.comment.replies)} 条回复 →",
                        font=ctk.CTkFont(size=10),
                        text_color=COLORS['secondary']).pack(anchor="w", padx=10, pady=5)
    
    def _show_reaction_picker(self):
        """显示表情选择器"""
        picker = ctk.CTkToplevel(self)
        picker.title("")
        picker.geometry("200x50")
        picker.overrideredirect(True)
        picker.transient(self)
        picker.lift()
        
        # 定位到按钮下方
        x = self.winfo_rootx() + 12
        y = self.winfo_rooty() + self.winfo_height() - 50
        picker.geometry(f"+{x}+{y}")
        
        frame = ctk.CTkFrame(picker, fg_color=COLORS['surface'],
                            corner_radius=12, border_width=1,
                            border_color=COLORS['border'])
        frame.pack(fill="both", expand=True, padx=2, pady=2)
        
        for reaction in REACTIONS:
            btn = ctk.CTkButton(frame, text=reaction['emoji'],
                               width=24, height=24,
                               corner_radius=12,
                               fg_color="transparent",
                               hover_color=COLORS['surface_hover'],
                               font=ctk.CTkFont(size=14),
                               command=lambda e=reaction['emoji']: self._add_reaction(e, picker))
            btn.pack(side="left", padx=1)
        
        # 点击外部关闭
        picker.bind("<FocusOut>", lambda e: picker.destroy())
        picker.focus_force()
    
    def _add_reaction(self, emoji: str, picker):
        """添加表情回复"""
        picker.destroy()
        if self.on_reaction:
            self.on_reaction(self.comment.id, emoji)
    
    def _show_reply_input(self):
        """显示回复输入框"""
        if self._reply_entry:
            return
        
        input_frame = ctk.CTkFrame(self, fg_color=COLORS['background'],
                                  corner_radius=8)
        input_frame.pack(fill="x", padx=12, pady=(0, 10))
        
        self._reply_entry = ctk.CTkEntry(input_frame, 
                                        placeholder_text="写回复...",
                                        height=36, corner_radius=18,
                                        border_width=0,
                                        fg_color=COLORS['surface'])
        self._reply_entry.pack(side="left", fill="x", expand=True, padx=5, pady=5)
        
        send_btn = ctk.CTkButton(input_frame, text="发送", width=50, height=30,
                                corner_radius=15,
                                fg_color=COLORS['primary'],
                                font=ctk.CTkFont(size=11),
                                command=self._send_reply)
        send_btn.pack(side="right", padx=5, pady=5)
        
        self._reply_entry.focus()
        self._reply_entry.bind("<Return>", lambda e: self._send_reply())
    
    def _send_reply(self):
        """发送回复"""
        if not self._reply_entry:
            return
        
        text = self._reply_entry.get().strip()
        if text and self.on_reply:
            self.on_reply(self.comment.thread_id, text)
        
        # 清理输入框
        if self._reply_entry:
            self._reply_entry.master.destroy()
            self._reply_entry = None
    
    def _on_resolve_click(self):
        """解决评论"""
        if self.on_resolve:
            self.on_resolve(self.comment.thread_id)


class CommentSidePanel:
    """评论侧边栏面板 - 腾讯文档风格"""
    
    def __init__(self, app, comment_manager, 
                 on_navigate: Callable = None,
                 current_user_id: str = None,
                 current_user_name: str = None):
        """
        Args:
            app: 应用实例
            comment_manager: CommentManager 实例
            on_navigate: 导航到评论位置的回调
            current_user_id: 当前用户 ID
            current_user_name: 当前用户名
        """
        self.app = app
        self.comment_manager = comment_manager
        self.on_navigate = on_navigate
        self.current_user_id = current_user_id or "user"
        self.current_user_name = current_user_name or "我"
        
        self.panel: Optional[ctk.CTkToplevel] = None
        self._scroll_frame = None
        self._filter_var = None
        self._filter = "all"  # all, unresolved, mine
        
    def show(self):
        """显示评论面板"""
        if not CTK_AVAILABLE:
            return
            
        if self.panel and self.panel.winfo_exists():
            self.panel.lift()
            return
        
        self.panel = ctk.CTkToplevel(self.app)
        self.panel.title("💬 评论")
        self.panel.geometry("380x600")
        self.panel.transient(self.app)
        self.panel.configure(fg_color=COLORS['background'])
        
        # 定位到右侧
        x = self.app.winfo_x() + self.app.winfo_width() - 400
        y = self.app.winfo_y() + 50
        self.panel.geometry(f"+{x}+{y}")
        
        self._create_ui()
        self._refresh_comments()
    
    def hide(self):
        """隐藏面板"""
        if self.panel and self.panel.winfo_exists():
            self.panel.destroy()
        self.panel = None
    
    def _create_ui(self):
        """创建面板 UI"""
        # 头部
        header = ctk.CTkFrame(self.panel, fg_color=COLORS['surface'],
                             height=60, corner_radius=0)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        ctk.CTkLabel(header, text="💬 评论",
                    font=ctk.CTkFont(size=16, weight="bold"),
                    text_color=COLORS['text']).pack(side="left", padx=20, pady=15)
        
        # 筛选下拉框
        self._filter_var = ctk.StringVar(value="全部")
        filter_menu = ctk.CTkOptionMenu(header, values=["全部", "未解决", "我的"],
                                       variable=self._filter_var,
                                       width=80, height=30,
                                       corner_radius=15,
                                       fg_color=COLORS['surface_hover'],
                                       button_color=COLORS['surface_hover'],
                                       button_hover_color=COLORS['border'],
                                       text_color=COLORS['text'],
                                       font=ctk.CTkFont(size=11),
                                       command=self._on_filter_change)
        filter_menu.pack(side="right", padx=20, pady=15)
        
        # 评论列表
        self._scroll_frame = ctk.CTkScrollableFrame(self.panel, 
                                                    fg_color=COLORS['background'])
        self._scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 底部：添加评论
        footer = ctk.CTkFrame(self.panel, fg_color=COLORS['surface'],
                             height=60, corner_radius=0)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)
        
        ctk.CTkLabel(footer, text="💡 选中文本后右键添加评论",
                    font=ctk.CTkFont(size=11),
                    text_color=COLORS['text_muted']).pack(expand=True)
    
    def _on_filter_change(self, value: str):
        """筛选变化"""
        filter_map = {"全部": "all", "未解决": "unresolved", "我的": "mine"}
        self._filter = filter_map.get(value, "all")
        self._refresh_comments()
    
    def _refresh_comments(self):
        """刷新评论列表"""
        if not self._scroll_frame:
            return
        
        # 清空现有内容
        for widget in self._scroll_frame.winfo_children():
            widget.destroy()
        
        # 获取评论
        threads = self.comment_manager.get_all_threads(
            include_resolved=(self._filter != "unresolved")
        )
        
        # 应用筛选
        if self._filter == "mine":
            threads = [t for t in threads 
                      if any(c.author_id == self.current_user_id 
                            for c in t.comments)]
        
        if not threads:
            ctk.CTkLabel(self._scroll_frame, text="暂无评论",
                        font=ctk.CTkFont(size=13),
                        text_color=COLORS['text_muted']).pack(pady=50)
            return
        
        # 创建评论气泡
        for thread in threads:
            if thread.comments:
                main_comment = thread.comments[0]
                
                # 转换为 CommentData
                comment_data = CommentData(
                    id=main_comment.id,
                    thread_id=thread.id,
                    author_id=main_comment.author_id,
                    author_name=main_comment.author_name,
                    author_color=getattr(main_comment, 'color', '#4ECDC4'),
                    content=main_comment.content,
                    created_at=main_comment.created_at,
                    resolved=thread.resolved,
                    reactions=[],
                    replies=[
                        CommentData(
                            id=r.id,
                            thread_id=thread.id,
                            author_id=r.author_id,
                            author_name=r.author_name,
                            author_color=getattr(r, 'color', '#45B7D1'),
                            content=r.content,
                            created_at=r.created_at,
                            resolved=False,
                            reactions=[],
                            replies=[]
                        ) for r in thread.comments[1:]
                    ]
                )
                
                bubble = CommentBubble(
                    self._scroll_frame,
                    comment_data,
                    on_reply=self._on_reply,
                    on_resolve=self._on_resolve,
                    on_reaction=self._on_reaction,
                    current_user_id=self.current_user_id
                )
                bubble.pack(fill="x", pady=5)
                
                # 点击导航
                bubble.bind("<Button-1>", 
                           lambda e, t=thread: self._navigate_to_comment(t))
    
    def _on_reply(self, thread_id: str, content: str):
        """回复评论"""
        self.comment_manager.add_reply(
            thread_id, content,
            self.current_user_id, self.current_user_name
        )
        self._refresh_comments()
    
    def _on_resolve(self, thread_id: str):
        """解决评论"""
        self.comment_manager.resolve_thread(thread_id)
        self._refresh_comments()
    
    def _on_reaction(self, comment_id: str, emoji: str):
        """添加表情回复"""
        # TODO: 实现表情回复存储
        pass
    
    def _navigate_to_comment(self, thread):
        """导航到评论位置"""
        if self.on_navigate and hasattr(thread, 'document_range'):
            start, end = thread.document_range
            self.on_navigate(start, end)
