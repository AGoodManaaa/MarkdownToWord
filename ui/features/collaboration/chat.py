# -*- coding: utf-8 -*-
"""协作聊天功能模块"""

import re
import time
from datetime import datetime
from typing import Optional, List, Dict, Callable

try:
    import customtkinter as ctk
except ImportError:
    ctk = None

from .theme import get_colors

# 导入图标系统
try:
    from ui.icons import EMOJI_PICKER
    # 合并所有表情到一个列表
    EMOJI_LIST = []
    for category, emojis in EMOJI_PICKER.items():
        EMOJI_LIST.extend(emojis[:8])  # 每个分类取前8个
except ImportError:
    # 常用表情符号（备用）
    EMOJI_LIST = [
        '😀', '😃', '😄', '😁', '😆', '😅', '🤣', '😂',
        '🙂', '😉', '😊', '😇', '🥰', '😍', '🤩', '😘',
        '😋', '😛', '😜', '🤪', '😝', '🤑', '🤗', '🤭',
        '🤔', '🤐', '😐', '😑', '😶', '😏', '😒', '🙄',
        '😬', '😮', '😯', '😲', '😳', '🥺', '😦', '😧',
        '😨', '😰', '😥', '😢', '😭', '😱', '😖', '😣',
        '👍', '👎', '👏', '🙌', '🤝', '🙏', '✌️', '🤞',
        '❤️', '🧡', '💛', '💚', '💙', '💜', '🖤', '🤍',
    ]


class ChatMessage:
    """聊天消息"""
    
    def __init__(self, sender_id: str, sender_name: str, content: str,
                 timestamp: float = None, mentions: List[str] = None,
                 color: str = None):
        self.sender_id = sender_id
        self.sender_name = sender_name
        self.content = content
        self.timestamp = timestamp or time.time()
        self.mentions = mentions or []
        self.color = color or '#4ECDC4'
    
    def to_dict(self) -> dict:
        return {
            'sender_id': self.sender_id,
            'sender_name': self.sender_name,
            'content': self.content,
            'timestamp': self.timestamp,
            'mentions': self.mentions,
            'color': self.color,
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'ChatMessage':
        return cls(
            sender_id=data.get('sender_id', ''),
            sender_name=data.get('sender_name', ''),
            content=data.get('content', ''),
            timestamp=data.get('timestamp'),
            mentions=data.get('mentions', []),
            color=data.get('color'),
        )
    
    @property
    def time_str(self) -> str:
        """格式化时间"""
        dt = datetime.fromtimestamp(self.timestamp)
        return dt.strftime("%H:%M")


class ChatPanel:
    """聊天面板"""
    
    def __init__(self, app, feature):
        self.app = app
        self.feature = feature
        self.panel: Optional[ctk.CTkFrame] = None
        self._messages: List[ChatMessage] = []
        self._message_frame = None
        self._input_entry = None
        self._on_send_callback: Optional[Callable] = None
        self._mention_popup = None
        self._participants: Dict[str, str] = {}  # name -> id
        self._user_name = "我"  # 当前用户名
        self._user_color = "#10b981"  # 当前用户颜色
    
    def set_user_info(self, name: str, color: str = None):
        """设置当前用户信息"""
        self._user_name = name or "我"
        if color:
            self._user_color = color
    
    def set_participants(self, participants: Dict[str, str]):
        """设置参与者列表（用于@提及）"""
        self._participants = participants
    
    def set_on_send(self, callback: Callable):
        """设置发送消息回调"""
        self._on_send_callback = callback
    
    def add_message(self, message: ChatMessage):
        """添加消息"""
        self._messages.append(message)
        self._render_message(message)
        self._scroll_to_bottom()
    
    def show(self, parent) -> ctk.CTkFrame:
        """显示聊天面板"""
        colors = get_colors()
        
        self.panel = ctk.CTkFrame(parent, fg_color=colors['surface'], corner_radius=0)
        
        # 标题栏
        header = ctk.CTkFrame(self.panel, fg_color=colors['background_secondary'], height=50)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        header_left = ctk.CTkFrame(header, fg_color="transparent")
        header_left.pack(side="left", padx=15, pady=10)
        
        ctk.CTkLabel(header_left, text="💬 聊天",
                    font=ctk.CTkFont(size=14, weight="bold"),
                    text_color=colors['text']).pack(side="left")
        
        # 在线人数
        online_count = len(self._participants)
        if online_count > 0:
            ctk.CTkLabel(header_left, text=f" · {online_count}人在线",
                        font=ctk.CTkFont(size=11),
                        text_color=colors['text_muted']).pack(side="left")
        
        # 消息区域
        self._message_frame = ctk.CTkScrollableFrame(
            self.panel, fg_color=colors['background'],
            corner_radius=0
        )
        self._message_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 渲染已有消息
        for msg in self._messages:
            self._render_message(msg)
        
        # 输入区域
        input_frame = ctk.CTkFrame(self.panel, fg_color=colors['background_secondary'], height=65)
        input_frame.pack(fill="x", side="bottom")
        input_frame.pack_propagate(False)
        
        # 表情按钮
        emoji_btn = ctk.CTkButton(
            input_frame, text="😊", width=40, height=40,
            fg_color="transparent", hover_color=colors['surface_hover'],
            corner_radius=20,
            command=self._show_emoji_picker
        )
        emoji_btn.pack(side="left", padx=8, pady=12)
        
        # 输入框
        self._input_entry = ctk.CTkEntry(
            input_frame, placeholder_text="输入消息，@提及他人...",
            height=40, corner_radius=20,
            fg_color=colors['surface'],
            border_color=colors['border']
        )
        self._input_entry.pack(side="left", fill="x", expand=True, padx=5, pady=12)
        self._input_entry.bind('<Return>', self._on_send)
        self._input_entry.bind('<KeyRelease>', self._on_key_release)
        
        # 发送按钮
        send_btn = ctk.CTkButton(
            input_frame, text="发送", width=70, height=40,
            corner_radius=20, fg_color=colors['primary'],
            hover_color=colors['primary_dark'],
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._on_send
        )
        send_btn.pack(side="right", padx=10, pady=12)
        
        return self.panel
    
    def _render_message(self, message: ChatMessage):
        """渲染单条消息"""
        if not self._message_frame:
            return
        
        colors = get_colors()
        is_self = message.sender_id == 'self'  # 判断是否是自己发的
        
        # 消息容器
        msg_frame = ctk.CTkFrame(self._message_frame, fg_color="transparent")
        msg_frame.pack(fill="x", pady=4, padx=5)
        
        if is_self:
            # 自己的消息靠右
            inner = ctk.CTkFrame(msg_frame, fg_color="transparent")
            inner.pack(side="right")
            
            content_frame = ctk.CTkFrame(inner, fg_color="transparent")
            content_frame.pack(side="right")
            
            # 发送者名称和时间（右对齐）
            name_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
            name_frame.pack(anchor="e")
            
            ctk.CTkLabel(name_frame, text=message.time_str,
                        font=ctk.CTkFont(size=9),
                        text_color=colors['text_muted']).pack(side="right")
            
            ctk.CTkLabel(name_frame, text=message.sender_name,
                        font=ctk.CTkFont(size=10, weight="bold"),
                        text_color=colors['primary']).pack(side="right", padx=5)
            
            # 消息气泡
            bubble = ctk.CTkFrame(content_frame, fg_color=colors['primary_light'], corner_radius=12)
            bubble.pack(anchor="e")
            
            ctk.CTkLabel(bubble, text=message.content,
                        font=ctk.CTkFont(size=12),
                        text_color=colors['text'],
                        wraplength=220).pack(padx=12, pady=8)
            
            # 头像
            avatar = ctk.CTkLabel(
                inner, text=message.sender_name[0].upper() if message.sender_name else "我",
                width=32, height=32, fg_color=message.color or self._user_color,
                corner_radius=16, font=ctk.CTkFont(size=12, weight="bold"),
                text_color="white"
            )
            avatar.pack(side="right", padx=(8, 0))
        else:
            # 他人的消息靠左
            inner = ctk.CTkFrame(msg_frame, fg_color="transparent")
            inner.pack(side="left")
            
            # 头像
            avatar = ctk.CTkLabel(
                inner, text=message.sender_name[0].upper() if message.sender_name else "?",
                width=32, height=32, fg_color=message.color or '#4ECDC4',
                corner_radius=16, font=ctk.CTkFont(size=12, weight="bold"),
                text_color="white"
            )
            avatar.pack(side="left", padx=(0, 8))
            
            content_frame = ctk.CTkFrame(inner, fg_color="transparent")
            content_frame.pack(side="left")
            
            # 发送者名称和时间
            name_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
            name_frame.pack(anchor="w")
            
            ctk.CTkLabel(name_frame, text=message.sender_name,
                        font=ctk.CTkFont(size=10, weight="bold"),
                        text_color=colors['text_secondary']).pack(side="left")
            
            ctk.CTkLabel(name_frame, text=message.time_str,
                        font=ctk.CTkFont(size=9),
                        text_color=colors['text_muted']).pack(side="left", padx=5)
            
            # 消息气泡
            bubble = ctk.CTkFrame(content_frame, fg_color=colors['surface_hover'], corner_radius=12)
            bubble.pack(anchor="w")
            
            # 处理@提及高亮
            content = self._highlight_mentions(message.content, colors)
            ctk.CTkLabel(bubble, text=content,
                        font=ctk.CTkFont(size=12),
                        text_color=colors['text'],
                        wraplength=220).pack(padx=12, pady=8)
    
    def _highlight_mentions(self, content: str, colors: dict) -> str:
        """高亮@提及（简单实现，返回原文本）"""
        # 实际高亮需要使用富文本，这里简化处理
        return content
    
    def _scroll_to_bottom(self):
        """滚动到底部"""
        if self._message_frame:
            self._message_frame._parent_canvas.yview_moveto(1.0)
    
    def _on_send(self, event=None):
        """发送消息"""
        if not self._input_entry:
            return
        
        content = self._input_entry.get().strip()
        if not content:
            return
        
        # 解析@提及
        mentions = re.findall(r'@(\w+)', content)
        
        # 创建消息（使用实际用户名）
        message = ChatMessage(
            sender_id='self',
            sender_name=self._user_name,
            content=content,
            mentions=mentions,
            color=self._user_color
        )
        
        # 添加到本地
        self.add_message(message)
        
        # 清空输入框
        self._input_entry.delete(0, 'end')
        
        # 回调发送
        if self._on_send_callback:
            self._on_send_callback(message)
    
    def _on_key_release(self, event):
        """按键释放事件（用于@提及）"""
        if not self._input_entry:
            return
        
        content = self._input_entry.get()
        cursor_pos = self._input_entry.index('insert')
        
        # 检查是否正在输入@
        if '@' in content:
            # 找到最近的@位置
            at_pos = content.rfind('@', 0, cursor_pos)
            if at_pos >= 0:
                query = content[at_pos+1:cursor_pos]
                if query and not ' ' in query:
                    self._show_mention_popup(query)
                    return
        
        self._hide_mention_popup()
    
    def _show_mention_popup(self, query: str):
        """显示@提及弹窗"""
        # 过滤匹配的参与者
        matches = [name for name in self._participants.keys() 
                   if query.lower() in name.lower()]
        
        if not matches:
            self._hide_mention_popup()
            return
        
        colors = get_colors()
        
        if self._mention_popup:
            self._hide_mention_popup()
        
        self._mention_popup = ctk.CTkFrame(
            self.panel, fg_color=colors['surface'],
            corner_radius=8, border_width=1, border_color=colors['border']
        )
        self._mention_popup.place(relx=0.1, rely=0.7, anchor="sw")
        
        for name in matches[:5]:
            btn = ctk.CTkButton(
                self._mention_popup, text=f"@{name}",
                fg_color="transparent", hover_color=colors['surface_hover'],
                anchor="w", height=30,
                command=lambda n=name: self._insert_mention(n)
            )
            btn.pack(fill="x", padx=5, pady=2)
    
    def _hide_mention_popup(self):
        """隐藏@提及弹窗"""
        if self._mention_popup:
            self._mention_popup.destroy()
            self._mention_popup = None
    
    def _insert_mention(self, name: str):
        """插入@提及"""
        if not self._input_entry:
            return
        
        content = self._input_entry.get()
        cursor_pos = self._input_entry.index('insert')
        
        # 找到@位置并替换
        at_pos = content.rfind('@', 0, cursor_pos)
        if at_pos >= 0:
            new_content = content[:at_pos] + f"@{name} " + content[cursor_pos:]
            self._input_entry.delete(0, 'end')
            self._input_entry.insert(0, new_content)
        
        self._hide_mention_popup()
    
    def _show_emoji_picker(self):
        """显示表情选择器"""
        colors = get_colors()
        
        popup = ctk.CTkToplevel(self.app)
        popup.title("😊 表情")
        popup.geometry("360x280")
        popup.resizable(False, False)
        popup.transient(self.app)
        popup.configure(fg_color=colors['background'])
        
        # 标题
        header = ctk.CTkFrame(popup, fg_color=colors['background_secondary'], height=40)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        ctk.CTkLabel(header, text="选择表情",
                    font=ctk.CTkFont(size=12, weight="bold"),
                    text_color=colors['text']).pack(side="left", padx=15, pady=10)
        
        # 表情网格
        frame = ctk.CTkScrollableFrame(popup, fg_color=colors['background'])
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        row_frame = None
        for i, emoji in enumerate(EMOJI_LIST):
            if i % 8 == 0:
                row_frame = ctk.CTkFrame(frame, fg_color="transparent")
                row_frame.pack(fill="x", pady=2)
            
            btn = ctk.CTkButton(
                row_frame, text=emoji, width=36, height=36,
                fg_color="transparent", hover_color=colors['surface_hover'],
                corner_radius=8,
                font=ctk.CTkFont(size=20),
                command=lambda e=emoji, p=popup: self._insert_emoji(e, p)
            )
            btn.pack(side="left", padx=3, pady=2)
    
    def _insert_emoji(self, emoji: str, popup):
        """插入表情"""
        if self._input_entry:
            self._input_entry.insert('insert', emoji)
        popup.destroy()
    
    def hide(self):
        """隐藏面板"""
        if self.panel:
            self.panel.destroy()
            self.panel = None
