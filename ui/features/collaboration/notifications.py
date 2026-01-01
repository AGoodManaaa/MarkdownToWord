# -*- coding: utf-8 -*-
"""协作通知系统 - 腾讯文档风格

Features:
- 桌面通知推送
- 音效提示
- Toast 通知动画
- 通知历史记录
"""

from dataclasses import dataclass, field
from typing import List, Optional, Callable, Dict, Any
from datetime import datetime
from enum import Enum
import threading
import queue

try:
    import customtkinter as ctk
    CTK_AVAILABLE = True
except ImportError:
    ctk = None
    CTK_AVAILABLE = False


class NotificationType(Enum):
    """通知类型"""
    USER_JOIN = "user_join"
    USER_LEAVE = "user_leave"
    COMMENT_NEW = "comment_new"
    COMMENT_REPLY = "comment_reply"
    MENTION = "mention"
    EDIT_CONFLICT = "edit_conflict"
    VERSION_SAVED = "version_saved"
    MEETING_END = "meeting_end"
    NETWORK_ISSUE = "network_issue"


# 通知图标和颜色配置
NOTIFICATION_CONFIG = {
    NotificationType.USER_JOIN: {
        "icon": "👋",
        "color": "#10b981",
        "sound": "join",
        "title": "用户加入"
    },
    NotificationType.USER_LEAVE: {
        "icon": "👋",
        "color": "#6b7280",
        "sound": None,
        "title": "用户离开"
    },
    NotificationType.COMMENT_NEW: {
        "icon": "💬",
        "color": "#3b82f6",
        "sound": "comment",
        "title": "新评论"
    },
    NotificationType.COMMENT_REPLY: {
        "icon": "💬",
        "color": "#3b82f6",
        "sound": "reply",
        "title": "新回复"
    },
    NotificationType.MENTION: {
        "icon": "@",
        "color": "#f59e0b",
        "sound": "mention",
        "title": "有人@你"
    },
    NotificationType.EDIT_CONFLICT: {
        "icon": "⚠️",
        "color": "#ef4444",
        "sound": "alert",
        "title": "编辑冲突"
    },
    NotificationType.VERSION_SAVED: {
        "icon": "✅",
        "color": "#10b981",
        "sound": None,
        "title": "版本已保存"
    },
    NotificationType.MEETING_END: {
        "icon": "🔚",
        "color": "#6b7280",
        "sound": "end",
        "title": "会议结束"
    },
    NotificationType.NETWORK_ISSUE: {
        "icon": "📶",
        "color": "#ef4444",
        "sound": "alert",
        "title": "网络问题"
    },
}


@dataclass
class Notification:
    """通知数据"""
    id: str
    type: NotificationType
    title: str
    message: str
    timestamp: datetime = field(default_factory=datetime.now)
    read: bool = False
    data: Dict = field(default_factory=dict)
    
    @property
    def config(self) -> Dict:
        """获取通知配置"""
        return NOTIFICATION_CONFIG.get(self.type, {})
    
    @property
    def icon(self) -> str:
        return self.config.get("icon", "📢")
    
    @property
    def color(self) -> str:
        return self.config.get("color", "#6b7280")
    
    @property
    def time_ago(self) -> str:
        """获取相对时间"""
        delta = datetime.now() - self.timestamp
        seconds = delta.total_seconds()
        
        if seconds < 60:
            return "刚刚"
        elif seconds < 3600:
            return f"{int(seconds / 60)} 分钟前"
        elif seconds < 86400:
            return f"{int(seconds / 3600)} 小时前"
        else:
            return self.timestamp.strftime("%m-%d %H:%M")


class NotificationToast:
    """通知 Toast 动画组件"""
    
    def __init__(self, parent, notification: Notification, 
                 on_click: Callable = None,
                 on_dismiss: Callable = None):
        self.parent = parent
        self.notification = notification
        self.on_click = on_click
        self.on_dismiss = on_dismiss
        self.frame: Optional[Any] = None
        self._animation_id = None
        
    def show(self, duration: int = 4000):
        """显示 Toast"""
        if not CTK_AVAILABLE:
            return
            
        colors = {
            'surface': '#ffffff',
            'text': '#1f2937',
            'text_secondary': '#6b7280',
            'border': '#e5e7eb',
        }
        
        # 创建 Toast 框架
        self.frame = ctk.CTkFrame(
            self.parent,
            fg_color=colors['surface'],
            corner_radius=12,
            border_width=1,
            border_color=self.notification.color
        )
        
        # 定位到右上角
        x = self.parent.winfo_width() - 340
        y = 20
        self.frame.place(x=x, y=y, width=320, height=80)
        
        # 左侧颜色条
        color_bar = ctk.CTkFrame(self.frame, fg_color=self.notification.color,
                                width=4, corner_radius=2)
        color_bar.pack(side="left", fill="y", padx=(8, 0), pady=8)
        
        # 内容区
        content = ctk.CTkFrame(self.frame, fg_color="transparent")
        content.pack(side="left", fill="both", expand=True, padx=12, pady=8)
        
        # 图标 + 标题行
        header = ctk.CTkFrame(content, fg_color="transparent")
        header.pack(fill="x")
        
        ctk.CTkLabel(header, text=self.notification.icon,
                    font=ctk.CTkFont(size=16)).pack(side="left")
        
        ctk.CTkLabel(header, text=self.notification.title,
                    font=ctk.CTkFont(size=12, weight="bold"),
                    text_color=colors['text']).pack(side="left", padx=8)
        
        ctk.CTkLabel(header, text=self.notification.time_ago,
                    font=ctk.CTkFont(size=10),
                    text_color=colors['text_secondary']).pack(side="right")
        
        # 消息内容
        ctk.CTkLabel(content, text=self.notification.message,
                    font=ctk.CTkFont(size=11),
                    text_color=colors['text_secondary'],
                    wraplength=260, justify="left").pack(anchor="w", pady=(4, 0))
        
        # 关闭按钮
        close_btn = ctk.CTkButton(self.frame, text="×", width=24, height=24,
                                 corner_radius=12,
                                 fg_color="transparent",
                                 hover_color=colors['border'],
                                 text_color=colors['text_secondary'],
                                 font=ctk.CTkFont(size=14),
                                 command=self._dismiss)
        close_btn.place(relx=1.0, rely=0, x=-30, y=5)
        
        # 点击处理
        self.frame.bind("<Button-1>", lambda e: self._on_click())
        
        # 自动消失
        self._animation_id = self.parent.after(duration, self._fade_out)
    
    def _on_click(self):
        """点击处理"""
        if self.on_click:
            self.on_click(self.notification)
        self._dismiss()
    
    def _dismiss(self):
        """关闭 Toast"""
        if self._animation_id:
            try:
                self.parent.after_cancel(self._animation_id)
            except Exception:
                pass
        
        if self.frame and self.frame.winfo_exists():
            self.frame.destroy()
        
        if self.on_dismiss:
            self.on_dismiss(self.notification)
    
    def _fade_out(self, step: int = 0):
        """淡出动画"""
        if not self.frame or not self.frame.winfo_exists():
            return
        
        max_steps = 10
        if step < max_steps:
            # 向上滑动
            try:
                x = self.parent.winfo_width() - 340
                y = 20 - step * 2
                self.frame.place(x=x, y=y)
                self._animation_id = self.parent.after(30, lambda: self._fade_out(step + 1))
            except Exception:
                self._dismiss()
        else:
            self._dismiss()


class NotificationManager:
    """通知管理器"""
    
    def __init__(self, app=None, enable_sound: bool = True):
        self.app = app
        self.enable_sound = enable_sound
        self._notifications: List[Notification] = []
        self._toast_queue: queue.Queue = queue.Queue()
        self._active_toast: Optional[NotificationToast] = None
        self._listeners: List[Callable] = []
        self._notification_id = 0
        self._max_history = 50
        
    def notify(self, type: NotificationType, message: str, 
               data: Dict = None, show_toast: bool = True) -> Notification:
        """发送通知
        
        Args:
            type: 通知类型
            message: 通知消息
            data: 附加数据
            show_toast: 是否显示 Toast
            
        Returns:
            Notification 对象
        """
        self._notification_id += 1
        config = NOTIFICATION_CONFIG.get(type, {})
        
        notification = Notification(
            id=f"notif_{self._notification_id}",
            type=type,
            title=config.get("title", "通知"),
            message=message,
            data=data or {}
        )
        
        # 添加到历史
        self._notifications.insert(0, notification)
        if len(self._notifications) > self._max_history:
            self._notifications = self._notifications[:self._max_history]
        
        # 播放声音
        if self.enable_sound and config.get("sound"):
            self._play_sound(config["sound"])
        
        # 显示 Toast
        if show_toast and self.app:
            self._show_toast(notification)
        
        # 通知监听器
        self._notify_listeners(notification)
        
        return notification
    
    def _show_toast(self, notification: Notification):
        """显示 Toast 通知"""
        if not self.app or not CTK_AVAILABLE:
            return
        
        def show():
            toast = NotificationToast(
                self.app, notification,
                on_click=self._on_toast_click,
                on_dismiss=self._on_toast_dismiss
            )
            self._active_toast = toast
            toast.show()
        
        if self.app:
            self.app.after(0, show)
    
    def _on_toast_click(self, notification: Notification):
        """Toast 点击回调"""
        notification.read = True
        self._notify_listeners(notification)
    
    def _on_toast_dismiss(self, notification: Notification):
        """Toast 关闭回调"""
        self._active_toast = None
        # 显示队列中的下一个
        if not self._toast_queue.empty():
            next_notif = self._toast_queue.get()
            self._show_toast(next_notif)
    
    def _play_sound(self, sound_name: str):
        """播放提示音"""
        # 使用 Windows 系统声音
        try:
            import winsound
            if sound_name == "join":
                winsound.MessageBeep(winsound.MB_OK)
            elif sound_name == "mention":
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
            elif sound_name == "alert":
                winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
            else:
                winsound.MessageBeep()
        except Exception:
            pass
    
    def add_listener(self, callback: Callable):
        """添加通知监听器"""
        if callback not in self._listeners:
            self._listeners.append(callback)
    
    def remove_listener(self, callback: Callable):
        """移除监听器"""
        if callback in self._listeners:
            self._listeners.remove(callback)
    
    def _notify_listeners(self, notification: Notification):
        """通知所有监听器"""
        for listener in self._listeners:
            try:
                listener(notification)
            except Exception:
                pass
    
    def get_unread_count(self) -> int:
        """获取未读数量"""
        return sum(1 for n in self._notifications if not n.read)
    
    def get_all(self, limit: int = 20) -> List[Notification]:
        """获取所有通知"""
        return self._notifications[:limit]
    
    def mark_all_read(self):
        """标记全部已读"""
        for n in self._notifications:
            n.read = True
    
    def clear_all(self):
        """清空通知"""
        self._notifications.clear()
    
    # 便捷方法
    def notify_user_join(self, user_name: str):
        """通知用户加入"""
        return self.notify(NotificationType.USER_JOIN, f"{user_name} 加入了会议")
    
    def notify_user_leave(self, user_name: str):
        """通知用户离开"""
        return self.notify(NotificationType.USER_LEAVE, f"{user_name} 离开了会议")
    
    def notify_mention(self, from_user: str, content: str):
        """通知@提及"""
        return self.notify(NotificationType.MENTION, f"{from_user} 提及了你: {content}")
    
    def notify_comment(self, from_user: str, content: str):
        """通知新评论"""
        return self.notify(NotificationType.COMMENT_NEW, f"{from_user}: {content}")
    
    def notify_conflict(self, user_name: str, location: str):
        """通知编辑冲突"""
        return self.notify(
            NotificationType.EDIT_CONFLICT,
            f"与 {user_name} 在 {location} 发生编辑冲突"
        )
    
    def notify_network_issue(self, message: str):
        """通知网络问题"""
        return self.notify(NotificationType.NETWORK_ISSUE, message)


class NotificationPanel:
    """通知面板 - 显示历史通知"""
    
    def __init__(self, app, notification_manager: NotificationManager):
        self.app = app
        self.manager = notification_manager
        self.panel: Optional[Any] = None
        self._scroll_frame = None
        
    def show(self):
        """显示通知面板"""
        if not CTK_AVAILABLE:
            return
            
        if self.panel and self.panel.winfo_exists():
            self.panel.lift()
            return
        
        colors = {
            'surface': '#ffffff',
            'background': '#f9fafb',
            'text': '#1f2937',
            'text_secondary': '#6b7280',
            'text_muted': '#9ca3af',
            'border': '#e5e7eb',
            'primary': '#10b981',
        }
        
        self.panel = ctk.CTkToplevel(self.app)
        self.panel.title("🔔 通知")
        self.panel.geometry("360x480")
        self.panel.transient(self.app)
        self.panel.configure(fg_color=colors['background'])
        
        # 头部
        header = ctk.CTkFrame(self.panel, fg_color=colors['surface'], height=50)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        unread_count = self.manager.get_unread_count()
        title_text = f"🔔 通知 ({unread_count})" if unread_count else "🔔 通知"
        
        ctk.CTkLabel(header, text=title_text,
                    font=ctk.CTkFont(size=14, weight="bold"),
                    text_color=colors['text']).pack(side="left", padx=20, pady=12)
        
        if unread_count:
            ctk.CTkButton(header, text="全部已读", width=70, height=28,
                         corner_radius=14,
                         fg_color=colors['primary'],
                         font=ctk.CTkFont(size=10),
                         command=self._mark_all_read).pack(side="right", padx=20, pady=12)
        
        # 通知列表
        self._scroll_frame = ctk.CTkScrollableFrame(self.panel, fg_color="transparent")
        self._scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self._refresh()
    
    def _refresh(self):
        """刷新通知列表"""
        if not self._scroll_frame:
            return
        
        colors = {
            'surface': '#ffffff',
            'surface_hover': '#f3f4f6',
            'text': '#1f2937',
            'text_secondary': '#6b7280',
            'text_muted': '#9ca3af',
            'border': '#e5e7eb',
        }
        
        # 清空
        for widget in self._scroll_frame.winfo_children():
            widget.destroy()
        
        notifications = self.manager.get_all()
        
        if not notifications:
            ctk.CTkLabel(self._scroll_frame, text="暂无通知",
                        font=ctk.CTkFont(size=13),
                        text_color=colors['text_muted']).pack(pady=50)
            return
        
        for notif in notifications:
            bg = colors['surface'] if notif.read else colors['surface_hover']
            
            frame = ctk.CTkFrame(self._scroll_frame, fg_color=bg, corner_radius=10)
            frame.pack(fill="x", pady=3)
            
            # 左侧颜色点
            if not notif.read:
                ctk.CTkLabel(frame, text="●", text_color=notif.color,
                            font=ctk.CTkFont(size=8)).pack(side="left", padx=(10, 0))
            
            # 图标
            ctk.CTkLabel(frame, text=notif.icon,
                        font=ctk.CTkFont(size=14)).pack(side="left", padx=8, pady=10)
            
            # 内容
            content_frame = ctk.CTkFrame(frame, fg_color="transparent")
            content_frame.pack(side="left", fill="x", expand=True, pady=8)
            
            # 标题行
            title_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
            title_frame.pack(fill="x")
            
            ctk.CTkLabel(title_frame, text=notif.title,
                        font=ctk.CTkFont(size=11, weight="bold"),
                        text_color=colors['text']).pack(side="left")
            
            ctk.CTkLabel(title_frame, text=notif.time_ago,
                        font=ctk.CTkFont(size=9),
                        text_color=colors['text_muted']).pack(side="right", padx=10)
            
            # 消息
            ctk.CTkLabel(content_frame, text=notif.message,
                        font=ctk.CTkFont(size=10),
                        text_color=colors['text_secondary'],
                        wraplength=240, justify="left").pack(anchor="w")
    
    def _mark_all_read(self):
        """标记全部已读"""
        self.manager.mark_all_read()
        self._refresh()
    
    def hide(self):
        """隐藏面板"""
        if self.panel and self.panel.winfo_exists():
            self.panel.destroy()
        self.panel = None
