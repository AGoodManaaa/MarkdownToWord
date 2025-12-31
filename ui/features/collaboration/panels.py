# -*- coding: utf-8 -*-
"""协作功能面板模块 - 完整优化版"""

import asyncio
import json
import os
import threading
import time
import tkinter as tk
from datetime import datetime
from tkinter import messagebox
from typing import Optional, List, Dict, Callable

try:
    import customtkinter as ctk
except ImportError:
    ctk = None

from .server import CollaborationServer
from .client import CollaborationClient
from .cursor import CursorManager
from .comments import CommentManager
from .history import HistoryManager
from .mentions import MentionManager
from .chat import ChatPanel, ChatMessage
from .version_history import VersionHistoryManager, VersionHistoryPanel
from .network import NetworkMonitor, NetworkStats, ConnectionQuality, OfflineCache, DataCompressor
from .security import PasswordChecker, Encryptor, AccessControl, TokenManager

# 导入图标系统
try:
    from ui.icons import icons, get_message_icon, EMOJI_PICKER
except ImportError:
    icons = None
    get_message_icon = lambda x: ""
    EMOJI_PICKER = {}

# 获取 app.ico 的绝对路径
APP_ICO_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'app.ico')

# 历史记录文件路径
HISTORY_FILE = os.path.join(os.path.dirname(__file__), 'meeting_history.json')

# 浅色配色方案（之前的好看配色）
COLORS = {
    'primary': '#10b981',
    'primary_dark': '#059669',
    'primary_light': '#d1fae5',
    'secondary': '#3b82f6',
    'secondary_dark': '#2563eb',
    'danger': '#ef4444',
    'danger_dark': '#dc2626',
    'warning': '#f59e0b',
    'gray': '#6b7280',
    'gray_light': '#f3f4f6',
    'text': '#1f2937',
    'text_secondary': '#6b7280',
    'text_muted': '#9ca3af',
    'surface': '#ffffff',
    'surface_hover': '#f3f4f6',
    'border': '#e5e7eb',
    'background': '#ffffff',
    'background_secondary': '#f9fafb',
    'toast_info': ('#3b82f6', '#dbeafe'),
    'toast_success': ('#10b981', '#d1fae5'),
    'toast_warning': ('#f59e0b', '#fef3c7'),
    'toast_error': ('#ef4444', '#fee2e2'),
}


def get_colors():
    """获取配色"""
    return COLORS


def set_window_icon(window):
    """设置窗口图标"""
    try:
        if os.path.exists(APP_ICO_PATH):
            window.after(100, lambda: window.iconbitmap(APP_ICO_PATH) if window.winfo_exists() else None)
    except Exception:
        pass


class AnimatedToast:
    """带动画的 Toast 通知"""
    
    def __init__(self, app):
        self.app = app
        self._toasts: List[ctk.CTkFrame] = []
        self._animation_steps = 10
        self._animation_delay = 30
    
    def show(self, message: str, type: str = 'info', duration: int = 3000):
        """显示带淡入淡出动画的 Toast"""
        colors = get_colors()
        toast_colors = {
            'info': colors['toast_info'],
            'success': colors['toast_success'],
            'warning': colors['toast_warning'],
            'error': colors['toast_error'],
        }
        fg_color, bg_color = toast_colors.get(type, colors['toast_info'])
        
        toast = ctk.CTkFrame(self.app, fg_color=bg_color, corner_radius=10)
        toast.place(relx=0.5, rely=-0.1, anchor="n")  # 从屏幕外开始
        
        icon = {'info': '💡', 'success': '✅', 'warning': '⚠️', 'error': '❌'}.get(type, '💡')
        ctk.CTkLabel(toast, text=f"{icon} {message}", font=ctk.CTkFont(size=12),
                    text_color=fg_color).pack(padx=20, pady=10)
        
        self._toasts.append(toast)
        
        # 淡入动画
        self._animate_in(toast, 0)
        
        # 延迟后淡出
        self.app.after(duration, lambda: self._animate_out(toast, 0))
    
    def _animate_in(self, toast, step):
        """淡入动画"""
        if step <= self._animation_steps:
            rely = -0.1 + (0.12 * step / self._animation_steps)
            toast.place(relx=0.5, rely=rely, anchor="n")
            self.app.after(self._animation_delay, lambda: self._animate_in(toast, step + 1))
    
    def _animate_out(self, toast, step):
        """淡出动画"""
        if not toast.winfo_exists():
            return
        
        if step <= self._animation_steps:
            rely = 0.02 - (0.12 * step / self._animation_steps)
            toast.place(relx=0.5, rely=rely, anchor="n")
            self.app.after(self._animation_delay, lambda: self._animate_out(toast, step + 1))
        else:
            if toast.winfo_exists():
                toast.destroy()
            if toast in self._toasts:
                self._toasts.remove(toast)


class MeetingHistory:
    """会议历史记录管理"""
    
    def __init__(self):
        self._history: List[Dict] = []
        self._load()
    
    def _load(self):
        try:
            if os.path.exists(HISTORY_FILE):
                with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                    self._history = json.load(f)
        except Exception:
            self._history = []
    
    def _save(self):
        try:
            with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
                json.dump(self._history[-20:], f, ensure_ascii=False, indent=2)
        except Exception:
            pass
    
    def add(self, address: str, code: str, name: str, is_host: bool = False):
        record = {
            'address': address, 'code': code, 'name': name,
            'is_host': is_host, 'timestamp': datetime.now().isoformat(),
        }
        self._history = [h for h in self._history if h.get('code') != code]
        self._history.append(record)
        self._save()
    
    def get_recent(self, limit: int = 5) -> List[Dict]:
        return list(reversed(self._history[-limit:]))
    
    def clear(self):
        self._history = []
        self._save()


class CollaborationFeature:
    """协作功能入口 - 完整优化版"""
    
    def __init__(self, app):
        self.app = app
        self.server: Optional[CollaborationServer] = None
        self.client: Optional[CollaborationClient] = None
        self.cursor_manager: Optional[CursorManager] = None
        self.comment_manager: Optional[CommentManager] = None
        self.history_manager: Optional[HistoryManager] = None
        self.mention_manager: Optional[MentionManager] = None
        
        # 新增管理器
        self.version_manager = VersionHistoryManager()
        self.network_monitor = NetworkMonitor()
        self.offline_cache = OfflineCache()
        self.access_control = AccessControl()
        self.encryptor: Optional[Encryptor] = None
        self.token_manager = TokenManager()
        
        # 状态
        self._is_host = False
        self._session_code: Optional[str] = None
        self._meeting_name: Optional[str] = None
        self._server_address: Optional[str] = None
        self._participants: List[Dict] = []
        self._dialog: Optional['CollaborationDialog'] = None
        self._status_bar: Optional['CollaborationStatusBar'] = None
        self._chat_panel: Optional[ChatPanel] = None
        self._server_loop = None
        self._server_thread = None
        self._start_time: Optional[float] = None
        self._reconnect_attempts = 0
        self._max_reconnect_attempts = 5
        self._heartbeat_interval = 30
        self._toast = AnimatedToast(app)
        self._meeting_history = MeetingHistory()
        self._user_name = "用户"
        
        # 设置网络监控回调
        self.network_monitor.set_on_quality_change(self._on_network_quality_change)
        self.network_monitor.set_on_disconnect(self._on_network_disconnect)
        
        # 注册快捷键
        self._register_shortcuts()
    
    def _register_shortcuts(self):
        try:
            self.app.bind('<Control-Shift-C>', lambda e: self.show_dialog())
        except Exception:
            pass
    
    def _on_network_quality_change(self, new_quality, old_quality):
        """网络质量变化回调"""
        if new_quality in (ConnectionQuality.POOR, ConnectionQuality.BAD):
            self.show_toast(f"网络质量{self.network_monitor.stats.quality_text}", 'warning')
        if self._status_bar:
            self._status_bar.update_network_status()
    
    def _on_network_disconnect(self):
        """网络断开回调"""
        self.show_toast("网络连接已断开", 'error')
        self.offline_cache.set_offline(True)
    
    def show_dialog(self) -> None:
        if ctk is None:
            messagebox.showerror("错误", "CustomTkinter 未安装")
            return
        if self._dialog is None:
            self._dialog = CollaborationDialog(self.app, self)
        self._dialog.show()
    
    def show_status_bar(self) -> None:
        if self._status_bar is None:
            self._status_bar = CollaborationStatusBar(self.app, self)
        self._status_bar.show()
    
    def hide_status_bar(self) -> None:
        if self._status_bar:
            self._status_bar.hide()
    
    def show_toast(self, message: str, type: str = 'info'):
        self._toast.show(message, type)
    
    def show_chat_panel(self, parent) -> ChatPanel:
        """显示聊天面板"""
        if self._chat_panel is None:
            self._chat_panel = ChatPanel(self.app, self)
            self._chat_panel.set_on_send(self._on_chat_send)
        
        # 设置用户信息
        user_name = self._user_name if not self._is_host else "主持人"
        self._chat_panel.set_user_info(user_name)
        
        return self._chat_panel.show(parent)
    
    def _on_chat_send(self, message: ChatMessage):
        """聊天消息发送回调"""
        if self.client:
            self.client.send_chat(message.to_dict())
    
    def start_hosting(self, document: str, password: str = None,
                      on_success: Callable = None, custom_code: str = None,
                      port: int = 8765, meeting_name: str = None,
                      enable_encryption: bool = False,
                      enable_waiting_room: bool = False) -> None:
        """开始主持协作会话"""
        self._is_host = True
        self._meeting_name = meeting_name or "未命名会议"
        self._start_time = time.time()
        
        # 设置加密
        if enable_encryption and password:
            self.encryptor = Encryptor(password)
        
        # 设置访问控制
        if enable_waiting_room:
            self.access_control.enable_waiting_room(True)
        
        self.server = CollaborationServer(port=port)
        
        def start_server():
            self._server_loop = asyncio.new_event_loop()
            asyncio.set_event_loop(self._server_loop)
            
            try:
                address = self._server_loop.run_until_complete(self.server.start())
                self._server_address = address
                
                self._session_code = self.server.create_session(
                    document, host_id="host", password=password,
                    custom_code=custom_code, meeting_name=meeting_name
                )
                
                self._meeting_history.add(address, self._session_code, self._meeting_name, is_host=True)
                
                if on_success:
                    self.app.after(0, lambda: on_success(self._session_code, address))
                
                self.app.after(0, self.show_status_bar)
                self.app.after(0, lambda: self.show_toast(f"会议 {self._meeting_name} 已创建", 'success'))
                
                # 保存初始版本
                self.version_manager.add_version(document, "host", "主持人", "初始版本")
                
                self._server_loop.run_forever()
                
            except Exception as e:
                error_msg = str(e)
                self.app.after(0, lambda msg=error_msg: messagebox.showerror("错误", f"启动服务器失败: {msg}"))
            finally:
                if self._server_loop:
                    self._server_loop.close()
        
        self._server_thread = threading.Thread(target=start_server, daemon=True)
        self._server_thread.start()
        self._init_managers()

    def join_session(self, address: str, session_code: str, name: str,
                     password: str = None, on_success: Callable = None,
                     on_error: Callable = None) -> None:
        """加入协作会话"""
        self._is_host = False
        self._session_code = session_code
        self._server_address = address
        self._start_time = time.time()
        self._reconnect_attempts = 0
        self._user_name = name
        
        # 设置加密
        if password:
            self.encryptor = Encryptor(password)
        
        self.client = CollaborationClient(self.app)
        self._setup_client_callbacks()
        
        def wrapped_success():
            self._meeting_history.add(address, session_code, name, is_host=False)
            self.app.after(0, self.show_status_bar)
            self.app.after(0, lambda: self.show_toast("已成功加入会议", 'success'))
            self.offline_cache.set_offline(False)
            if on_success:
                on_success()
        
        def wrapped_error(msg):
            self.app.after(0, lambda: self.show_toast(f"连接失败: {msg}", 'error'))
            if on_error:
                on_error(msg)
        
        self.client.connect(address=address, session_code=session_code, name=name,
                           password=password, on_success=wrapped_success, on_error=wrapped_error)
        self._init_managers()
        self._start_heartbeat()
        self._start_ping()

    def _start_heartbeat(self):
        def heartbeat():
            if self.is_active and self.client:
                try:
                    if hasattr(self.client, 'send_heartbeat'):
                        self.client.send_heartbeat()
                except Exception:
                    self._handle_disconnect()
                self.app.after(self._heartbeat_interval * 1000, heartbeat)
        self.app.after(self._heartbeat_interval * 1000, heartbeat)
    
    def _start_ping(self):
        """启动 ping 检测"""
        def ping():
            if self.is_active and self.client:
                start = time.time()
                try:
                    if hasattr(self.client, 'ping'):
                        self.client.ping()
                        latency = (time.time() - start) * 1000
                        self.network_monitor.record_ping(latency)
                        if self._status_bar:
                            self._status_bar.update_network_status()
                except Exception:
                    pass
                self.app.after(5000, ping)
        self.app.after(5000, ping)
    
    def _handle_disconnect(self):
        if self._reconnect_attempts < self._max_reconnect_attempts:
            self._reconnect_attempts += 1
            self.show_toast(f"连接断开，正在重连 ({self._reconnect_attempts}/{self._max_reconnect_attempts})...", 'warning')
            self.offline_cache.set_offline(True)
            self._try_reconnect()
        else:
            self.show_toast("连接已断开，无法重连", 'error')
            self.disconnect()
    
    def _try_reconnect(self):
        if self.client and self._server_address and self._session_code:
            def on_success():
                self._reconnect_attempts = 0
                self.offline_cache.set_offline(False)
                self.show_toast("重连成功", 'success')
                self._sync_offline_changes()
            
            def on_error(msg):
                self.app.after(3000, self._handle_disconnect)
            
            self.client.connect(
                address=self._server_address,
                session_code=self._session_code,
                name=self._user_name,
                on_success=on_success,
                on_error=on_error
            )
    
    def _sync_offline_changes(self):
        """同步离线更改"""
        if self.offline_cache.has_pending():
            operations = self.offline_cache.get_pending_operations()
            for op in operations:
                if self.client:
                    self.client.send_operation(op)
            self.offline_cache.clear_operations()
            self.show_toast(f"已同步 {len(operations)} 个离线更改", 'success')

    def _init_managers(self) -> None:
        editor = getattr(self.app, 'input_text', None)
        self.cursor_manager = CursorManager(editor)
        self.comment_manager = CommentManager(self.app)
        self.history_manager = HistoryManager()
        self.mention_manager = MentionManager(self.app)
    
    def _setup_client_callbacks(self) -> None:
        if not self.client:
            return
        self.client.on_remote_operation(self._on_remote_operation)
        self.client.on_cursor_update(self._on_cursor_update)
        self.client.on_participant_change(self._on_participant_change)
        self.client.on_comment(self._on_comment)
        self.client.on_kicked(self._on_kicked)
        
        # 聊天回调
        if hasattr(self.client, 'on_chat'):
            self.client.on_chat(self._on_chat_receive)
    
    def _on_remote_operation(self, operation: dict) -> None:
        if not self.client:
            return
        content = self.client.get_document()
        
        # 解密
        if self.encryptor and self.encryptor.is_enabled:
            content = self.encryptor.decrypt(content)
        
        if hasattr(self.app, 'input_text'):
            try:
                cursor_pos = self.app.input_text._textbox.index("insert")
            except:
                cursor_pos = "1.0"
            self.app.input_text.delete("1.0", "end")
            self.app.input_text.insert("1.0", content)
            try:
                self.app.input_text._textbox.mark_set("insert", cursor_pos)
            except:
                pass
    
    def _on_cursor_update(self, participant_id: str, position: int, selection: tuple = None) -> None:
        if self.cursor_manager:
            self.cursor_manager.update_cursor(participant_id, position, selection)

    def _on_participant_change(self, participants: list) -> None:
        old_ids = {p['id'] for p in self._participants if p.get('is_online')}
        new_ids = {p['id'] for p in participants if p.get('is_online')}
        
        for p in participants:
            if p['id'] in new_ids - old_ids and p.get('is_online'):
                self.show_toast(f"👋 {p['name']} 加入了会议", 'info')
        
        for p in self._participants:
            if p['id'] in old_ids - new_ids:
                self.show_toast(f"👋 {p['name']} 离开了会议", 'info')
        
        self._participants = participants
        
        if self.cursor_manager:
            current_ids = set(self.cursor_manager.cursors.keys())
            for p in participants:
                if p['id'] not in current_ids and p.get('is_online'):
                    self.cursor_manager.add_cursor(p['id'], p['name'], p.get('color'))
            for pid in current_ids - new_ids:
                self.cursor_manager.remove_cursor(pid)
        
        if self.mention_manager:
            mapping = {p['name']: p['id'] for p in participants}
            self.mention_manager.set_participants(mapping)
        
        if self._chat_panel:
            self._chat_panel.set_participants(mapping)
        
        if self._status_bar:
            self._status_bar.update_participants(participants)
        if self._dialog:
            self._dialog.update_participants(participants)
    
    def _on_comment(self, comment: dict) -> None:
        pass
    
    def _on_kicked(self, message: str) -> None:
        self.disconnect()
        self.show_toast(message, 'warning')
        messagebox.showwarning("已断开", message)
    
    def _on_chat_receive(self, data: dict):
        """接收聊天消息"""
        if self._chat_panel:
            message = ChatMessage.from_dict(data)
            self._chat_panel.add_message(message)
    
    def kick_participant(self, participant_id: str) -> bool:
        if not self._is_host or not self.server:
            return False
        try:
            if hasattr(self.server, 'kick_participant'):
                self.server.kick_participant(participant_id)
                self.show_toast("已移除参与者", 'success')
                return True
        except Exception:
            pass
        return False
    
    def set_participant_permission(self, participant_id: str, permission: str) -> bool:
        if not self._is_host or not self.server:
            return False
        try:
            if hasattr(self.server, 'set_permission'):
                self.server.set_permission(participant_id, permission)
                self.show_toast(f"已更新权限", 'success')
                return True
        except Exception:
            pass
        return False
    
    def lock_meeting(self):
        """锁定会议"""
        self.access_control.lock_meeting()
        self.show_toast("会议已锁定", 'info')
    
    def unlock_meeting(self):
        """解锁会议"""
        self.access_control.unlock_meeting()
        self.show_toast("会议已解锁", 'info')
    
    def disconnect(self) -> None:
        if self.client:
            self.client.disconnect()
            self.client = None
        if self.cursor_manager:
            self.cursor_manager.clear_all()
        self._session_code = None
        self._participants = []
        self._start_time = None
        self.network_monitor.reset()
        self.hide_status_bar()
    
    def stop_hosting(self) -> None:
        if self._server_loop:
            self._server_loop.call_soon_threadsafe(self._server_loop.stop)
        self.server = None
        self.disconnect()
        self._is_host = False
        self._meeting_name = None
        self._server_address = None
        self.show_toast("会议已结束", 'info')
    
    def get_connection_duration(self) -> str:
        if not self._start_time:
            return "00:00"
        duration = int(time.time() - self._start_time)
        hours, remainder = divmod(duration, 3600)
        minutes, seconds = divmod(remainder, 60)
        if hours > 0:
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        return f"{minutes:02d}:{seconds:02d}"
    
    def get_invite_text(self) -> str:
        return f"📋 邀请加入协作\n会议名称: {self._meeting_name}\n会议码: {self._session_code}\n服务器地址: {self._server_address}"
    
    @property
    def is_active(self) -> bool:
        return self._is_host or (self.client is not None and self.client.is_connected)
    
    @property
    def is_connected(self) -> bool:
        return self.client is not None and self.client.is_connected
    
    @property
    def session_code(self) -> Optional[str]:
        return self._session_code
    
    @property
    def participant_count(self) -> int:
        return len([p for p in self._participants if p.get('is_online')])


class CollaborationStatusBar:
    """协作状态栏 - 现代化设计"""
    
    def __init__(self, app, feature: CollaborationFeature):
        self.app = app
        self.feature = feature
        self.frame: Optional[ctk.CTkFrame] = None
        self._outer_frame: Optional[ctk.CTkFrame] = None
        self._avatar_frame = None
        self._status_label = None
        self._time_label = None
        self._network_label = None
        self._sync_label = None
        self._pulse_state = True
        self._pulse_job = None
        self._time_job = None
    
    def show(self) -> None:
        if self.frame and self.frame.winfo_exists():
            return
        
        colors = get_colors()
        
        self._outer_frame = ctk.CTkFrame(self.app, fg_color="transparent", height=50)
        try:
            self._outer_frame.pack(side="top", fill="x", before=self.app.winfo_children()[0])
        except:
            self._outer_frame.pack(side="top", fill="x")
        
        # 增加宽度到 900px 确保按钮显示完整
        self.frame = ctk.CTkFrame(
            self._outer_frame, height=42, width=900,
            fg_color=colors['primary_light'], corner_radius=21,
            border_width=1, border_color=colors['primary']
        )
        self.frame.pack(pady=4)
        self.frame.pack_propagate(False)
        
        # 左侧：状态
        left_frame = ctk.CTkFrame(self.frame, fg_color="transparent")
        left_frame.pack(side="left", padx=15, pady=5)
        
        self._status_dot = ctk.CTkLabel(left_frame, text="●", text_color=colors['primary'],
                                        font=ctk.CTkFont(size=14))
        self._status_dot.pack(side="left")
        self._start_pulse()
        
        role_text = "🎯 主持中" if self.feature._is_host else "✨ 协作中"
        self._status_label = ctk.CTkLabel(left_frame, text=f" {role_text}",
                                          font=ctk.CTkFont(size=12, weight="bold"),
                                          text_color=colors['text'])
        self._status_label.pack(side="left", padx=5)
        
        # 会议名称
        ctk.CTkLabel(left_frame, text=f"| {self.feature._meeting_name or '会议'}",
                    font=ctk.CTkFont(size=11), text_color=colors['text_secondary']).pack(side="left", padx=5)
        
        # 时长
        self._time_label = ctk.CTkLabel(left_frame, text="⏱ 00:00",
                                        font=ctk.CTkFont(size=11), text_color=colors['text_secondary'])
        self._time_label.pack(side="left", padx=8)
        self._update_time()
        
        # 网络状态
        self._network_label = ctk.CTkLabel(left_frame, text="🟢",
                                           font=ctk.CTkFont(size=10), text_color=colors['text_muted'])
        self._network_label.pack(side="left", padx=3)
        
        # 头像
        self._avatar_frame = ctk.CTkFrame(self.frame, fg_color="transparent")
        self._avatar_frame.pack(side="left", padx=10)
        
        # 右侧按钮
        right_frame = ctk.CTkFrame(self.frame, fg_color="transparent")
        right_frame.pack(side="right", padx=10, pady=5)
        
        # 聊天按钮
        ctk.CTkButton(right_frame, text="💬", width=36, height=28,
                     command=self._show_chat_panel, fg_color=colors['gray'],
                     hover_color=colors['text_secondary'], corner_radius=14).pack(side="left", padx=3)
        
        # 版本历史按钮
        ctk.CTkButton(right_frame, text="📜", width=36, height=28,
                     command=self._show_version_history, fg_color=colors['gray'],
                     hover_color=colors['text_secondary'], corner_radius=14).pack(side="left", padx=3)
        
        # 邀请按钮
        ctk.CTkButton(right_frame, text="📤 邀请", width=70, height=28,
                     command=self._show_invite_dialog, fg_color=colors['secondary'],
                     hover_color=colors['secondary_dark'], corner_radius=14,
                     font=ctk.CTkFont(size=11)).pack(side="left", padx=3)
        
        # 参与者按钮
        ctk.CTkButton(right_frame, text="👥", width=36, height=28,
                     command=self._show_members_panel, fg_color=colors['gray'],
                     hover_color=colors['text_secondary'], corner_radius=14).pack(side="left", padx=3)
        
        # 设置按钮（主持人）
        if self.feature._is_host:
            ctk.CTkButton(right_frame, text="⚙️", width=36, height=28,
                         command=self._show_settings, fg_color=colors['gray'],
                         hover_color=colors['text_secondary'], corner_radius=14).pack(side="left", padx=3)
        
        # 结束/退出按钮
        if self.feature._is_host:
            ctk.CTkButton(right_frame, text="结束会议", width=80, height=28,
                         command=self._end_meeting, fg_color=colors['danger'],
                         hover_color=colors['danger_dark'], corner_radius=14,
                         font=ctk.CTkFont(size=11)).pack(side="left", padx=3)
        else:
            ctk.CTkButton(right_frame, text="退出", width=60, height=28,
                         command=self._leave_meeting, fg_color=colors['gray'],
                         hover_color=colors['text_secondary'], corner_radius=14,
                         font=ctk.CTkFont(size=11)).pack(side="left", padx=3)
    
    def _start_pulse(self):
        colors = get_colors()
        def pulse():
            if not self._status_dot or not self._status_dot.winfo_exists():
                return
            self._pulse_state = not self._pulse_state
            color = colors['primary'] if self._pulse_state else colors['primary_dark']
            self._status_dot.configure(text_color=color)
            self._pulse_job = self.app.after(800, pulse)
        pulse()
    
    def _update_time(self):
        if self._time_label and self._time_label.winfo_exists():
            self._time_label.configure(text=f"⏱ {self.feature.get_connection_duration()}")
            self._time_job = self.app.after(1000, self._update_time)
    
    def update_network_status(self):
        """更新网络状态显示"""
        if self._network_label and self._network_label.winfo_exists():
            stats = self.feature.network_monitor.stats
            self._network_label.configure(text=f"{stats.quality_icon}")

    def hide(self) -> None:
        if self._pulse_job:
            self.app.after_cancel(self._pulse_job)
        if self._time_job:
            self.app.after_cancel(self._time_job)
        if self._outer_frame and self._outer_frame.winfo_exists():
            self._outer_frame.destroy()
        self._outer_frame = None
        self.frame = None
    
    def update_participants(self, participants: list) -> None:
        if not self._avatar_frame or not self._avatar_frame.winfo_exists():
            return
        
        for widget in self._avatar_frame.winfo_children():
            widget.destroy()
        
        online = [p for p in participants if p.get('is_online')][:5]
        
        for i, p in enumerate(online):
            avatar = ctk.CTkLabel(self._avatar_frame, text=p['name'][0].upper(),
                                 width=28, height=28, fg_color=p.get('color', '#4ECDC4'),
                                 corner_radius=14, font=ctk.CTkFont(size=11, weight="bold"),
                                 text_color="white")
            avatar.pack(side="left", padx=(-8 if i > 0 else 0, 0))
        
        colors = get_colors()
        if len(participants) > 5:
            ctk.CTkLabel(self._avatar_frame, text=f"+{len(participants)-5}",
                        width=28, height=28, fg_color=colors['gray_light'],
                        corner_radius=14, font=ctk.CTkFont(size=10),
                        text_color=colors['text_secondary']).pack(side="left", padx=(-8, 0))
    
    def _show_invite_dialog(self) -> None:
        colors = get_colors()
        
        dialog = ctk.CTkToplevel(self.app)
        dialog.title("邀请协作")
        dialog.geometry("420x400")
        dialog.resizable(False, False)
        dialog.transient(self.app)
        set_window_icon(dialog)
        dialog.lift()
        dialog.focus_force()
        
        ctk.CTkLabel(dialog, text="📤 邀请他人加入",
                    font=ctk.CTkFont(size=18, weight="bold"),
                    text_color=colors['text']).pack(pady=20)
        
        card = ctk.CTkFrame(dialog, fg_color=colors['surface'], corner_radius=12)
        card.pack(pady=10, padx=25, fill="x")
        
        ctk.CTkLabel(card, text="会议码", font=ctk.CTkFont(size=11),
                    text_color=colors['text_secondary']).pack(pady=(15,5))
        ctk.CTkLabel(card, text=self.feature._session_code or "",
                    font=ctk.CTkFont(size=28, weight="bold"),
                    text_color=colors['primary']).pack()
        
        ctk.CTkLabel(card, text="服务器地址", font=ctk.CTkFont(size=11),
                    text_color=colors['text_secondary']).pack(pady=(15,5))
        ctk.CTkLabel(card, text=self.feature._server_address or "",
                    font=ctk.CTkFont(size=12), text_color=colors['text']).pack(pady=(0,15))
        
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=20)
        
        def copy_code():
            dialog.clipboard_clear()
            dialog.clipboard_append(self.feature._session_code or "")
            self.feature.show_toast("会议码已复制", 'success')
        
        def copy_all():
            dialog.clipboard_clear()
            dialog.clipboard_append(self.feature.get_invite_text())
            self.feature.show_toast("邀请信息已复制", 'success')
        
        ctk.CTkButton(btn_frame, text="📋 复制会议码", command=copy_code,
                     width=130, height=36, corner_radius=18,
                     fg_color=colors['secondary']).pack(side="left", padx=5)
        
        ctk.CTkButton(btn_frame, text="📝 复制全部", command=copy_all,
                     width=130, height=36, corner_radius=18,
                     fg_color=colors['primary']).pack(side="left", padx=5)
        
        ctk.CTkLabel(dialog, text="💡 分享会议码和地址给其他人即可加入",
                    font=ctk.CTkFont(size=11), text_color=colors['text_secondary']).pack(pady=10)

    def _show_members_panel(self) -> None:
        colors = get_colors()
        
        dialog = ctk.CTkToplevel(self.app)
        dialog.title("参与者")
        dialog.geometry("360x480")
        dialog.transient(self.app)
        set_window_icon(dialog)
        dialog.lift()
        dialog.focus_force()
        
        header = ctk.CTkFrame(dialog, fg_color=colors['background_secondary'], height=50)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        ctk.CTkLabel(header, text=f"👥 参与者 ({self.feature.participant_count}人在线)",
                    font=ctk.CTkFont(size=14, weight="bold"),
                    text_color=colors['text']).pack(side="left", padx=20, pady=12)
        
        scroll = ctk.CTkScrollableFrame(dialog, fg_color=colors['background'])
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        for p in self.feature._participants:
            if p.get('is_online'):
                self._create_participant_card(scroll, p, dialog, colors)
    
    def _create_participant_card(self, parent, participant: dict, dialog, colors: dict) -> None:
        frame = ctk.CTkFrame(parent, fg_color=colors['surface'], corner_radius=10)
        frame.pack(fill="x", pady=4)
        
        avatar = ctk.CTkLabel(frame, text=participant['name'][0].upper(),
                             width=36, height=36, fg_color=participant.get('color', '#4ECDC4'),
                             corner_radius=18, font=ctk.CTkFont(size=14, weight="bold"),
                             text_color="white")
        avatar.pack(side="left", padx=10, pady=8)
        
        info_frame = ctk.CTkFrame(frame, fg_color="transparent")
        info_frame.pack(side="left", fill="x", expand=True, pady=8)
        
        name_frame = ctk.CTkFrame(info_frame, fg_color="transparent")
        name_frame.pack(anchor="w")
        
        ctk.CTkLabel(name_frame, text=participant['name'],
                    font=ctk.CTkFont(size=12, weight="bold"),
                    text_color=colors['text']).pack(side="left")
        
        is_host = participant.get('id') == 'host' or participant.get('is_host')
        if is_host:
            ctk.CTkLabel(name_frame, text="主持人", fg_color=colors['warning'],
                        corner_radius=4, font=ctk.CTkFont(size=9),
                        text_color="white", width=45, height=18).pack(side="left", padx=5)
        
        perm = participant.get('permission', 'edit')
        perm_text = {'edit': '可编辑', 'comment': '可评论', 'view': '仅查看'}.get(perm, perm)
        ctk.CTkLabel(info_frame, text=perm_text, font=ctk.CTkFont(size=10),
                    text_color=colors['text_secondary']).pack(anchor="w")
        
        if self.feature._is_host and not is_host:
            btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
            btn_frame.pack(side="right", padx=10)
            
            ctk.CTkButton(btn_frame, text="✕", width=28, height=28,
                         fg_color=colors['danger'], hover_color=colors['danger_dark'],
                         corner_radius=14, font=ctk.CTkFont(size=10),
                         command=lambda pid=participant['id']: self._kick_participant(pid, dialog)).pack()
    
    def _kick_participant(self, participant_id: str, dialog):
        if messagebox.askyesno("确认", "确定要移除该参与者吗？"):
            self.feature.kick_participant(participant_id)
            dialog.destroy()
            self._show_members_panel()
    
    def _show_chat_panel(self):
        """显示聊天面板"""
        colors = get_colors()
        
        dialog = ctk.CTkToplevel(self.app)
        dialog.title("💬 聊天")
        dialog.geometry("450x600")
        dialog.transient(self.app)
        dialog.configure(fg_color=colors['background'])
        set_window_icon(dialog)
        dialog.lift()
        dialog.focus_force()
        
        # 使用 ChatPanel 显示聊天
        chat_panel = self.feature.show_chat_panel(dialog)
        chat_panel.pack(fill="both", expand=True)
    
    def _show_version_history(self):
        """显示版本历史"""
        panel = VersionHistoryPanel(self.app, self.feature)
        
        def on_restore(content):
            if hasattr(self.app, 'input_text'):
                self.app.input_text.delete("1.0", "end")
                self.app.input_text.insert("1.0", content)
                self.feature.show_toast("已恢复到历史版本", 'success')
        
        panel.set_on_restore(on_restore)
        panel.show()
    
    def _show_settings(self):
        """显示设置面板"""
        colors = get_colors()
        
        dialog = ctk.CTkToplevel(self.app)
        dialog.title("会议设置")
        dialog.geometry("400x300")
        dialog.transient(self.app)
        set_window_icon(dialog)
        dialog.lift()
        
        ctk.CTkLabel(dialog, text="⚙️ 会议设置",
                    font=ctk.CTkFont(size=16, weight="bold"),
                    text_color=colors['text']).pack(pady=20)
        
        # 锁定会议
        lock_frame = ctk.CTkFrame(dialog, fg_color=colors['surface'], corner_radius=10)
        lock_frame.pack(fill="x", padx=20, pady=5)
        
        ctk.CTkLabel(lock_frame, text="🔒 锁定会议",
                    font=ctk.CTkFont(size=12), text_color=colors['text']).pack(side="left", padx=15, pady=12)
        ctk.CTkLabel(lock_frame, text="禁止新成员加入",
                    font=ctk.CTkFont(size=10), text_color=colors['text_secondary']).pack(side="left")
        
        lock_switch = ctk.CTkSwitch(lock_frame, text="", width=40,
                                    command=lambda: self._toggle_lock(lock_switch))
        lock_switch.pack(side="right", padx=15)
        if self.feature.access_control.is_locked:
            lock_switch.select()
        
        # 等候室
        waiting_frame = ctk.CTkFrame(dialog, fg_color=colors['surface'], corner_radius=10)
        waiting_frame.pack(fill="x", padx=20, pady=5)
        
        ctk.CTkLabel(waiting_frame, text="🚪 等候室",
                    font=ctk.CTkFont(size=12), text_color=colors['text']).pack(side="left", padx=15, pady=12)
        ctk.CTkLabel(waiting_frame, text="新成员需审批",
                    font=ctk.CTkFont(size=10), text_color=colors['text_secondary']).pack(side="left")
        
        waiting_switch = ctk.CTkSwitch(waiting_frame, text="", width=40,
                                       command=lambda: self._toggle_waiting_room(waiting_switch))
        waiting_switch.pack(side="right", padx=15)
        if self.feature.access_control.waiting_room_enabled:
            waiting_switch.select()
    
    def _toggle_lock(self, switch):
        if switch.get():
            self.feature.lock_meeting()
        else:
            self.feature.unlock_meeting()
    
    def _toggle_waiting_room(self, switch):
        self.feature.access_control.enable_waiting_room(switch.get())
        self.feature.show_toast("等候室已" + ("启用" if switch.get() else "禁用"), 'info')
    
    def _end_meeting(self) -> None:
        if messagebox.askyesno("结束会议", "确定要结束会议吗？\n所有参与者将被断开连接。"):
            self.feature.stop_hosting()
    
    def _leave_meeting(self) -> None:
        if messagebox.askyesno("退出会议", "确定要退出会议吗？"):
            self.feature.disconnect()


class CollaborationDialog:
    """协作对话框 - 现代卡片式设计"""
    
    def __init__(self, app, feature: CollaborationFeature):
        self.app = app
        self.feature = feature
        self.dialog: Optional[ctk.CTkToplevel] = None
        self._participants_frame = None
        self._loading_label = None
    
    def show(self) -> None:
        if self.dialog is not None and self.dialog.winfo_exists():
            self.dialog.lift()
            self.dialog.focus_force()
            return
        
        colors = get_colors()
        
        self.dialog = ctk.CTkToplevel(self.app)
        self.dialog.title("实时协作")
        self.dialog.geometry("560x720")
        self.dialog.minsize(520, 680)
        self.dialog.transient(self.app)
        self.dialog.configure(fg_color=colors['background'])
        set_window_icon(self.dialog)
        self.dialog.lift()
        self.dialog.focus_force()
        
        if self.feature.is_active:
            self._create_active_ui()
        else:
            self._create_connect_ui()
    
    def _create_connect_ui(self) -> None:
        colors = get_colors()
        
        # 标题
        header = ctk.CTkFrame(self.dialog, fg_color="transparent")
        header.pack(fill="x", padx=30, pady=20)
        
        ctk.CTkLabel(header, text="👥", font=ctk.CTkFont(size=40)).pack()
        ctk.CTkLabel(header, text="实时协作",
                    font=ctk.CTkFont(size=24, weight="bold"),
                    text_color=colors['text']).pack(pady=5)
        ctk.CTkLabel(header, text="与他人一起编辑文档",
                    font=ctk.CTkFont(size=12), text_color=colors['text_secondary']).pack()
        
        # 标签页
        tabview = ctk.CTkTabview(self.dialog, fg_color=colors['surface'])
        tabview.pack(fill="both", expand=True, padx=25, pady=10)
        
        self._create_host_tab(tabview.add("🚀 创建会议"))
        self._create_join_tab(tabview.add("🔗 加入会议"))
        self._create_history_tab(tabview.add("📋 历史"))
    
    def _create_host_tab(self, parent) -> None:
        colors = get_colors()
        
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 会议名称
        self._create_input_group(scroll, "会议名称", "meeting_name", "例如：项目讨论", colors)
        
        # 会议码
        self._create_input_group(scroll, "会议码（留空自动生成）", "custom_code", "例如：ABC123", colors)
        
        # 端口和密码
        row = ctk.CTkFrame(scroll, fg_color="transparent")
        row.pack(fill="x", pady=5)
        
        port_frame = ctk.CTkFrame(row, fg_color="transparent")
        port_frame.pack(side="left", fill="x", expand=True, padx=(0,5))
        ctk.CTkLabel(port_frame, text="端口号", font=ctk.CTkFont(size=11),
                    text_color=colors['text_secondary']).pack(anchor="w")
        self.host_port = ctk.CTkEntry(port_frame, height=40, corner_radius=10,
                                      fg_color=colors['surface'], border_color=colors['border'])
        self.host_port.insert(0, "8765")
        self.host_port.pack(fill="x", pady=3)
        
        pwd_frame = ctk.CTkFrame(row, fg_color="transparent")
        pwd_frame.pack(side="left", fill="x", expand=True, padx=(5,0))
        ctk.CTkLabel(pwd_frame, text="密码（可选）", font=ctk.CTkFont(size=11),
                    text_color=colors['text_secondary']).pack(anchor="w")
        self.host_password = ctk.CTkEntry(pwd_frame, show="*", height=40, corner_radius=10,
                                          fg_color=colors['surface'], border_color=colors['border'])
        self.host_password.pack(fill="x", pady=3)
        
        # 高级选项
        adv_frame = ctk.CTkFrame(scroll, fg_color=colors['surface'], corner_radius=10)
        adv_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(adv_frame, text="高级选项", font=ctk.CTkFont(size=11, weight="bold"),
                    text_color=colors['text']).pack(anchor="w", padx=15, pady=(10,5))
        
        # 加密
        enc_frame = ctk.CTkFrame(adv_frame, fg_color="transparent")
        enc_frame.pack(fill="x", padx=15, pady=3)
        ctk.CTkLabel(enc_frame, text="🔐 端到端加密", font=ctk.CTkFont(size=11),
                    text_color=colors['text_secondary']).pack(side="left")
        self.enable_encryption = ctk.CTkSwitch(enc_frame, text="", width=40)
        self.enable_encryption.pack(side="right")
        
        # 等候室
        wait_frame = ctk.CTkFrame(adv_frame, fg_color="transparent")
        wait_frame.pack(fill="x", padx=15, pady=(3,10))
        ctk.CTkLabel(wait_frame, text="🚪 启用等候室", font=ctk.CTkFont(size=11),
                    text_color=colors['text_secondary']).pack(side="left")
        self.enable_waiting_room = ctk.CTkSwitch(wait_frame, text="", width=40)
        self.enable_waiting_room.pack(side="right")
        
        # 创建按钮
        ctk.CTkButton(scroll, text="🚀 创建会议", command=self._on_start_hosting,
                     height=48, corner_radius=24, fg_color=colors['primary'],
                     hover_color=colors['primary_dark'],
                     font=ctk.CTkFont(size=14, weight="bold")).pack(fill="x", pady=20)
    
    def _create_input_group(self, parent, label: str, attr_name: str, placeholder: str, colors: dict):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(frame, text=label, font=ctk.CTkFont(size=11),
                    text_color=colors['text_secondary']).pack(anchor="w")
        
        entry = ctk.CTkEntry(frame, placeholder_text=placeholder, height=40, corner_radius=10,
                            fg_color=colors['surface'], border_color=colors['border'])
        entry.pack(fill="x", pady=3)
        setattr(self, attr_name, entry)

    def _create_join_tab(self, parent) -> None:
        colors = get_colors()
        
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=5, pady=5)
        
        self._create_input_group(scroll, "服务器地址", "server_address", "ws://192.168.x.x:8765", colors)
        self.server_address.insert(0, "ws://localhost:8765")
        
        self._create_input_group(scroll, "会议码", "session_code", "输入会议码", colors)
        
        row = ctk.CTkFrame(scroll, fg_color="transparent")
        row.pack(fill="x", pady=5)
        
        name_frame = ctk.CTkFrame(row, fg_color="transparent")
        name_frame.pack(side="left", fill="x", expand=True, padx=(0,5))
        ctk.CTkLabel(name_frame, text="您的昵称", font=ctk.CTkFont(size=11),
                    text_color=colors['text_secondary']).pack(anchor="w")
        self.user_name = ctk.CTkEntry(name_frame, height=40, corner_radius=10,
                                      fg_color=colors['surface'], border_color=colors['border'])
        self.user_name.insert(0, "用户")
        self.user_name.pack(fill="x", pady=3)
        
        pwd_frame = ctk.CTkFrame(row, fg_color="transparent")
        pwd_frame.pack(side="left", fill="x", expand=True, padx=(5,0))
        ctk.CTkLabel(pwd_frame, text="密码（如需要）", font=ctk.CTkFont(size=11),
                    text_color=colors['text_secondary']).pack(anchor="w")
        self.join_password = ctk.CTkEntry(pwd_frame, show="*", height=40, corner_radius=10,
                                          fg_color=colors['surface'], border_color=colors['border'])
        self.join_password.pack(fill="x", pady=3)
        
        ctk.CTkButton(scroll, text="🔗 加入会议", command=self._on_join_session,
                     height=48, corner_radius=24, fg_color=colors['secondary'],
                     hover_color=colors['secondary_dark'],
                     font=ctk.CTkFont(size=14, weight="bold")).pack(fill="x", pady=20)
    
    def _create_history_tab(self, parent) -> None:
        colors = get_colors()
        
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=5, pady=5)
        
        history = self.feature._meeting_history.get_recent(10)
        
        if not history:
            ctk.CTkLabel(scroll, text="📭 暂无历史记录",
                        font=ctk.CTkFont(size=14), text_color=colors['text_secondary']).pack(pady=50)
            return
        
        for record in history:
            self._create_history_card(scroll, record, colors)
        
        ctk.CTkButton(scroll, text="🗑️ 清空历史", command=self._clear_history,
                     height=36, corner_radius=18, fg_color=colors['gray'],
                     hover_color=colors['text_secondary'],
                     font=ctk.CTkFont(size=12)).pack(pady=20)
    
    def _create_history_card(self, parent, record: dict, colors: dict) -> None:
        card = ctk.CTkFrame(parent, fg_color=colors['surface'], corner_radius=10)
        card.pack(fill="x", pady=4)
        
        info = ctk.CTkFrame(card, fg_color="transparent")
        info.pack(side="left", fill="x", expand=True, padx=12, pady=10)
        
        header = ctk.CTkFrame(info, fg_color="transparent")
        header.pack(anchor="w")
        
        ctk.CTkLabel(header, text=record.get('name', '未命名'),
                    font=ctk.CTkFont(size=12, weight="bold"),
                    text_color=colors['text']).pack(side="left")
        
        if record.get('is_host'):
            ctk.CTkLabel(header, text="主持", fg_color=colors['warning'],
                        corner_radius=4, font=ctk.CTkFont(size=9),
                        text_color="white", width=35, height=16).pack(side="left", padx=5)
        
        ctk.CTkLabel(info, text=f"会议码: {record.get('code', '')}",
                    font=ctk.CTkFont(size=10), text_color=colors['text_secondary']).pack(anchor="w")
        
        try:
            dt = datetime.fromisoformat(record.get('timestamp', ''))
            time_str = dt.strftime("%m-%d %H:%M")
        except:
            time_str = ""
        ctk.CTkLabel(info, text=time_str, font=ctk.CTkFont(size=9),
                    text_color=colors['text_muted']).pack(anchor="w")
        
        if not record.get('is_host'):
            ctk.CTkButton(card, text="加入", width=60, height=28, corner_radius=14,
                         fg_color=colors['secondary'], font=ctk.CTkFont(size=11),
                         command=lambda r=record: self._quick_join(r)).pack(side="right", padx=12)
    
    def _quick_join(self, record: dict) -> None:
        self.server_address.delete(0, "end")
        self.server_address.insert(0, record.get('address', ''))
        self.session_code.delete(0, "end")
        self.session_code.insert(0, record.get('code', ''))
        self.feature.show_toast("已填充会议信息", 'info')
    
    def _clear_history(self) -> None:
        if messagebox.askyesno("确认", "确定要清空所有历史记录吗？"):
            self.feature._meeting_history.clear()
            self.feature.show_toast("历史记录已清空", 'success')
            self._refresh_ui()

    def _create_active_ui(self) -> None:
        colors = get_colors()
        
        # 状态卡片
        status_card = ctk.CTkFrame(self.dialog, fg_color=colors['primary_light'], corner_radius=15)
        status_card.pack(fill="x", padx=25, pady=20)
        
        role = "🎯 主持中" if self.feature._is_host else "✨ 协作中"
        ctk.CTkLabel(status_card, text=role,
                    font=ctk.CTkFont(size=18, weight="bold"),
                    text_color=colors['primary']).pack(pady=(15,5))
        ctk.CTkLabel(status_card, text=self.feature._meeting_name or "会议",
                    font=ctk.CTkFont(size=14), text_color=colors['text']).pack()
        
        # 网络状态
        stats = self.feature.network_monitor.stats
        ctk.CTkLabel(status_card, text=f"⏱ {self.feature.get_connection_duration()} | {stats.quality_icon} {stats.latency_str}",
                    font=ctk.CTkFont(size=12), text_color=colors['text_secondary']).pack(pady=(5,15))
        
        # 会议信息
        info_card = ctk.CTkFrame(self.dialog, fg_color=colors['surface'], corner_radius=12)
        info_card.pack(fill="x", padx=25, pady=10)
        
        ctk.CTkLabel(info_card, text="会议码", font=ctk.CTkFont(size=11),
                    text_color=colors['text_secondary']).pack(pady=(12,3))
        ctk.CTkLabel(info_card, text=self.feature._session_code,
                    font=ctk.CTkFont(size=24, weight="bold"),
                    text_color=colors['primary']).pack()
        
        btn_frame = ctk.CTkFrame(info_card, fg_color="transparent")
        btn_frame.pack(pady=12)
        
        def copy_code():
            self.dialog.clipboard_clear()
            self.dialog.clipboard_append(self.feature._session_code or "")
            self.feature.show_toast("会议码已复制", 'success')
        
        def copy_all():
            self.dialog.clipboard_clear()
            self.dialog.clipboard_append(self.feature.get_invite_text())
            self.feature.show_toast("邀请信息已复制", 'success')
        
        ctk.CTkButton(btn_frame, text="📋 复制", command=copy_code,
                     width=100, height=34, corner_radius=17,
                     fg_color=colors['gray']).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="📤 邀请", command=copy_all,
                     width=100, height=34, corner_radius=17,
                     fg_color=colors['secondary']).pack(side="left", padx=5)
        
        # 参与者
        ctk.CTkLabel(self.dialog, text=f"👥 在线参与者 ({self.feature.participant_count})",
                    font=ctk.CTkFont(size=13, weight="bold"),
                    text_color=colors['text']).pack(pady=(15,8))
        
        self._participants_frame = ctk.CTkScrollableFrame(self.dialog, height=140,
                                                         fg_color="transparent")
        self._participants_frame.pack(fill="x", padx=25)
        self.update_participants(self.feature._participants)
        
        # 操作按钮
        btn_frame2 = ctk.CTkFrame(self.dialog, fg_color="transparent")
        btn_frame2.pack(pady=20)
        
        if self.feature._is_host:
            ctk.CTkButton(btn_frame2, text="🛑 结束会议", command=self._end_meeting,
                         width=160, height=44, corner_radius=22,
                         fg_color=colors['danger'], hover_color=colors['danger_dark'],
                         font=ctk.CTkFont(size=13)).pack()
        else:
            ctk.CTkButton(btn_frame2, text="🚪 退出会议", command=self._leave_meeting,
                         width=160, height=44, corner_radius=22,
                         fg_color=colors['gray'], hover_color=colors['text_secondary'],
                         font=ctk.CTkFont(size=13)).pack()

    def _on_start_hosting(self) -> None:
        password = self.host_password.get() or None
        meeting_name = self.meeting_name.get() or "未命名会议"
        custom_code = self.custom_code.get().strip().upper() or None
        enable_encryption = self.enable_encryption.get() if hasattr(self, 'enable_encryption') else False
        enable_waiting_room = self.enable_waiting_room.get() if hasattr(self, 'enable_waiting_room') else False
        
        try:
            port = int(self.host_port.get())
        except ValueError:
            port = 8765
        
        content = ""
        if hasattr(self.app, 'input_text'):
            content = self.app.input_text.get("1.0", "end-1c")
        
        def on_success(code, address):
            self._show_success_dialog(code, address, meeting_name)
        
        self.feature.start_hosting(content, password, on_success, custom_code, port, meeting_name,
                                   enable_encryption, enable_waiting_room)
        
        if self.dialog:
            self.dialog.destroy()
            self.dialog = None
    
    def _show_success_dialog(self, code: str, address: str, meeting_name: str) -> None:
        colors = get_colors()
        
        try:
            dialog = ctk.CTkToplevel(self.app)
            dialog.title("会议已创建")
            dialog.geometry("440x420")
            dialog.resizable(False, False)
            dialog.transient(self.app)
            dialog.configure(fg_color=colors['background'])
            set_window_icon(dialog)
            dialog.lift()
            
            ctk.CTkLabel(dialog, text="🎉", font=ctk.CTkFont(size=56)).pack(pady=(25,10))
            ctk.CTkLabel(dialog, text="会议创建成功！",
                        font=ctk.CTkFont(size=18, weight="bold"),
                        text_color=colors['text']).pack()
            ctk.CTkLabel(dialog, text=meeting_name,
                        font=ctk.CTkFont(size=13), text_color=colors['text_secondary']).pack(pady=5)
            
            card = ctk.CTkFrame(dialog, fg_color=colors['surface'], corner_radius=12)
            card.pack(pady=15, padx=30, fill="x")
            
            ctk.CTkLabel(card, text="会议码", font=ctk.CTkFont(size=11),
                        text_color=colors['text_secondary']).pack(pady=(12,3))
            ctk.CTkLabel(card, text=code, font=ctk.CTkFont(size=28, weight="bold"),
                        text_color=colors['primary']).pack()
            
            ctk.CTkLabel(card, text="服务器地址", font=ctk.CTkFont(size=11),
                        text_color=colors['text_secondary']).pack(pady=(10,3))
            ctk.CTkLabel(card, text=address, font=ctk.CTkFont(size=12),
                        text_color=colors['text']).pack(pady=(0,12))
            
            btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
            btn_frame.pack(pady=15)
            
            def copy_code():
                dialog.clipboard_clear()
                dialog.clipboard_append(code)
                self.feature.show_toast("会议码已复制", 'success')
            
            def copy_all():
                dialog.clipboard_clear()
                dialog.clipboard_append(self.feature.get_invite_text())
                self.feature.show_toast("邀请信息已复制", 'success')
            
            ctk.CTkButton(btn_frame, text="📋 复制会议码", command=copy_code,
                         width=130, height=40, corner_radius=20,
                         fg_color=colors['gray']).pack(side="left", padx=5)
            ctk.CTkButton(btn_frame, text="📝 复制全部", command=copy_all,
                         width=130, height=40, corner_radius=20,
                         fg_color=colors['primary']).pack(side="left", padx=5)
            
            ctk.CTkLabel(dialog, text="💡 分享给其他人即可加入协作",
                        font=ctk.CTkFont(size=11), text_color=colors['text_secondary']).pack(pady=10)
            
        except Exception:
            messagebox.showinfo("会议已创建", f"会议码: {code}\n地址: {address}")

    def _on_join_session(self) -> None:
        address = self.server_address.get()
        code = self.session_code.get()
        name = self.user_name.get() or "用户"
        password = self.join_password.get() or None
        
        if not address or not code:
            self.feature.show_toast("请填写服务器地址和会议码", 'warning')
            return
        
        def on_success():
            if self.dialog:
                self.dialog.destroy()
                self.dialog = None
        
        def on_error(msg):
            pass
        
        self.feature.join_session(address, code, name, password, on_success, on_error)
    
    def _end_meeting(self) -> None:
        if messagebox.askyesno("结束会议", "确定要结束会议吗？"):
            self.feature.stop_hosting()
            self._refresh_ui()
    
    def _leave_meeting(self) -> None:
        if messagebox.askyesno("退出会议", "确定要退出吗？"):
            self.feature.disconnect()
            self._refresh_ui()
    
    def _refresh_ui(self) -> None:
        if self.dialog and self.dialog.winfo_exists():
            for widget in self.dialog.winfo_children():
                widget.destroy()
            if self.feature.is_active:
                self._create_active_ui()
            else:
                self._create_connect_ui()
    
    def update_participants(self, participants: list) -> None:
        if not self._participants_frame or not self._participants_frame.winfo_exists():
            return
        
        colors = get_colors()
        
        for widget in self._participants_frame.winfo_children():
            widget.destroy()
        
        for p in participants:
            if p.get('is_online'):
                frame = ctk.CTkFrame(self._participants_frame, fg_color=colors['surface'],
                                    corner_radius=8)
                frame.pack(fill="x", pady=3)
                
                avatar = ctk.CTkLabel(frame, text=p['name'][0].upper(),
                                     width=30, height=30, fg_color=p.get('color', '#4ECDC4'),
                                     corner_radius=15, font=ctk.CTkFont(size=12, weight="bold"),
                                     text_color="white")
                avatar.pack(side="left", padx=8, pady=6)
                
                ctk.CTkLabel(frame, text=p.get('name', 'Unknown'),
                            font=ctk.CTkFont(size=12), text_color=colors['text']).pack(side="left", padx=5)
                
                is_host = p.get('id') == 'host' or p.get('is_host')
                if is_host:
                    ctk.CTkLabel(frame, text="主持人", fg_color=colors['warning'],
                                corner_radius=4, font=ctk.CTkFont(size=9),
                                text_color="white", width=45, height=18).pack(side="left", padx=5)
                
                perm = p.get('permission', 'edit')
                perm_text = {'edit': '编辑', 'comment': '评论', 'view': '查看'}.get(perm, perm)
                ctk.CTkLabel(frame, text=perm_text, font=ctk.CTkFont(size=10),
                            text_color=colors['text_secondary']).pack(side="right", padx=10)
