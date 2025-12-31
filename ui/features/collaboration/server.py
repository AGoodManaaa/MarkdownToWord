# -*- coding: utf-8 -*-
"""协作服务器模块"""

from dataclasses import dataclass, field
from typing import Dict, Set, Optional, Any
import asyncio
import json
import secrets
import hashlib
import time

try:
    import websockets
    from websockets.server import WebSocketServerProtocol
    WEBSOCKETS_AVAILABLE = True
except ImportError:
    WEBSOCKETS_AVAILABLE = False
    websockets = None
    WebSocketServerProtocol = None


@dataclass
class Participant:
    """参与者信息"""
    id: str
    name: str
    color: str
    cursor_position: int = 0
    selection: Optional[tuple] = None
    permission: str = 'edit'  # 'edit', 'comment', 'view'
    is_online: bool = True
    last_active: float = 0.0
    
    def to_dict(self) -> dict:
        """转换为字典"""
        return {
            'id': self.id,
            'name': self.name,
            'color': self.color,
            'cursor_position': self.cursor_position,
            'selection': self.selection,
            'permission': self.permission,
            'is_online': self.is_online
        }


@dataclass
class Session:
    """协作会话"""
    id: str
    host_id: str
    document_content: str
    participants: Dict[str, Participant] = field(default_factory=dict)
    password_hash: Optional[str] = None
    created_at: float = 0.0
    crdt_state: bytes = b''
    meeting_name: str = "未命名会议"
    
    def to_dict(self) -> dict:
        """转换为字典"""
        return {
            'id': self.id,
            'host_id': self.host_id,
            'participant_count': len(self.participants),
            'created_at': self.created_at,
            'meeting_name': self.meeting_name
        }


class CollaborationServer:
    """协作服务器（WebSocket）"""
    
    CURSOR_COLORS = [
        "#FF6B6B", "#4ECDC4", "#45B7D1", "#96CEB4",
        "#FFEAA7", "#DDA0DD", "#98D8C8", "#F7DC6F"
    ]
    
    def __init__(self, host: str = "0.0.0.0", port: int = 8765):
        """初始化协作服务器
        
        Args:
            host: 监听地址
            port: 监听端口
        """
        self.host = host
        self.port = port
        self.sessions: Dict[str, Session] = {}
        self.connections: Dict[str, 'WebSocketServerProtocol'] = {}
        self.participant_sessions: Dict[str, str] = {}  # participant_id -> session_id
        self._server = None
        self._color_index = 0
    
    async def start(self) -> str:
        """启动服务器
        
        Returns:
            连接地址
        """
        if not WEBSOCKETS_AVAILABLE:
            raise RuntimeError("websockets 库未安装")
        
        self._server = await websockets.serve(
            self.handle_connection,
            self.host,
            self.port
        )
        
        # 获取本机 IP 地址用于显示
        import socket
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            local_ip = s.getsockname()[0]
            s.close()
        except Exception:
            local_ip = "localhost"
        
        return f"ws://{local_ip}:{self.port}"
    
    async def stop(self) -> None:
        """停止服务器"""
        if self._server:
            self._server.close()
            await self._server.wait_closed()
            self._server = None
    
    def create_session(self, document: str, host_id: str, 
                       password: Optional[str] = None,
                       custom_code: Optional[str] = None,
                       meeting_name: Optional[str] = None) -> str:
        """创建新会话
        
        Args:
            document: 初始文档内容
            host_id: 主持人 ID
            password: 会话密码（可选）
            custom_code: 自定义会话码（可选）
            meeting_name: 会议名称（可选）
            
        Returns:
            会话码
        """
        # 使用自定义会话码或自动生成
        if custom_code and custom_code not in self.sessions:
            session_code = custom_code
        else:
            session_code = self._generate_session_code()
        
        password_hash = None
        if password:
            password_hash = hashlib.sha256(password.encode()).hexdigest()
        
        session = Session(
            id=session_code,
            host_id=host_id,
            document_content=document,
            password_hash=password_hash,
            created_at=time.time(),
            meeting_name=meeting_name or "未命名会议"
        )
        
        self.sessions[session_code] = session
        
        return session_code
    
    async def handle_connection(self, websocket: 'WebSocketServerProtocol') -> None:
        """处理 WebSocket 连接
        
        Args:
            websocket: WebSocket 连接
        """
        participant_id = None
        session_id = None
        
        try:
            async for message in websocket:
                data = json.loads(message)
                msg_type = data.get('type')
                
                if msg_type == 'join':
                    # 加入会话
                    result = await self._handle_join(websocket, data)
                    if result:
                        participant_id, session_id = result
                        self.connections[participant_id] = websocket
                        self.participant_sessions[participant_id] = session_id
                
                elif msg_type == 'leave':
                    # 离开会话
                    if participant_id and session_id:
                        await self._handle_leave(participant_id, session_id)
                
                elif msg_type == 'operation':
                    # 编辑操作
                    if session_id:
                        await self._handle_operation(participant_id, session_id, data)
                
                elif msg_type == 'cursor_update':
                    # 光标更新
                    if session_id:
                        await self._handle_cursor_update(participant_id, session_id, data)
                
                elif msg_type == 'sync_request':
                    # 同步请求
                    if session_id:
                        await self._handle_sync_request(websocket, session_id)
                
                elif msg_type == 'comment':
                    # 评论
                    if session_id:
                        await self._handle_comment(participant_id, session_id, data)
                
                elif msg_type == 'kick':
                    # 踢出用户
                    if session_id:
                        await self._handle_kick(participant_id, session_id, data)
        
        except Exception:
            pass
        finally:
            # 清理连接
            if participant_id:
                if participant_id in self.connections:
                    del self.connections[participant_id]
                if participant_id in self.participant_sessions:
                    session_id = self.participant_sessions[participant_id]
                    del self.participant_sessions[participant_id]
                    
                    if session_id in self.sessions:
                        session = self.sessions[session_id]
                        if participant_id in session.participants:
                            session.participants[participant_id].is_online = False
                            try:
                                await self._broadcast_presence(session_id)
                            except Exception:
                                pass
    
    async def _handle_join(self, websocket: 'WebSocketServerProtocol', 
                           data: dict) -> Optional[tuple]:
        """处理加入会话请求"""
        session_code = data.get('session_code')
        name = data.get('name', 'Anonymous')
        password = data.get('password')
        
        if session_code not in self.sessions:
            await websocket.send(json.dumps({
                'type': 'error',
                'message': '会话不存在'
            }))
            return None
        
        session = self.sessions[session_code]
        
        # 验证密码
        if session.password_hash:
            if not password:
                await websocket.send(json.dumps({
                    'type': 'error',
                    'message': '需要密码'
                }))
                return None
            
            if hashlib.sha256(password.encode()).hexdigest() != session.password_hash:
                await websocket.send(json.dumps({
                    'type': 'error',
                    'message': '密码错误'
                }))
                return None
        
        # 创建参与者
        participant_id = str(secrets.token_hex(8))
        color = self.CURSOR_COLORS[self._color_index % len(self.CURSOR_COLORS)]
        self._color_index += 1
        
        participant = Participant(
            id=participant_id,
            name=name,
            color=color,
            last_active=time.time()
        )
        
        session.participants[participant_id] = participant
        
        # 发送加入成功消息
        await websocket.send(json.dumps({
            'type': 'join_success',
            'participant_id': participant_id,
            'session_id': session_code,
            'color': color,
            'document': session.document_content,
            'participants': [p.to_dict() for p in session.participants.values()]
        }))
        
        # 广播新参与者
        await self._broadcast_presence(session_code)
        
        return (participant_id, session_code)
    
    async def _handle_leave(self, participant_id: str, session_id: str) -> None:
        """处理离开会话"""
        if session_id in self.sessions:
            session = self.sessions[session_id]
            if participant_id in session.participants:
                del session.participants[participant_id]
                await self._broadcast_presence(session_id)
    
    async def _handle_operation(self, participant_id: str, session_id: str, 
                                data: dict) -> None:
        """处理编辑操作"""
        if session_id not in self.sessions:
            return
        
        session = self.sessions[session_id]
        
        # 检查权限
        participant = session.participants.get(participant_id)
        if not participant or participant.permission == 'view':
            return
        
        # 更新文档内容
        operation = data.get('operation', {})
        
        # 广播操作给其他参与者
        await self.broadcast(session_id, {
            'type': 'operation',
            'operation': operation,
            'author': participant_id
        }, exclude=participant_id)
    
    async def _handle_cursor_update(self, participant_id: str, session_id: str,
                                    data: dict) -> None:
        """处理光标更新"""
        if session_id not in self.sessions:
            return
        
        session = self.sessions[session_id]
        participant = session.participants.get(participant_id)
        
        if participant:
            participant.cursor_position = data.get('position', 0)
            participant.selection = data.get('selection')
            participant.last_active = time.time()
            
            # 广播光标更新
            await self.broadcast(session_id, {
                'type': 'cursor_update',
                'participant_id': participant_id,
                'position': participant.cursor_position,
                'selection': participant.selection
            }, exclude=participant_id)
    
    async def _handle_sync_request(self, websocket: 'WebSocketServerProtocol',
                                   session_id: str) -> None:
        """处理同步请求"""
        if session_id not in self.sessions:
            return
        
        session = self.sessions[session_id]
        
        await websocket.send(json.dumps({
            'type': 'sync_response',
            'document': session.document_content,
            'crdt_state': session.crdt_state.hex() if session.crdt_state else ''
        }))
    
    async def _handle_comment(self, participant_id: str, session_id: str,
                              data: dict) -> None:
        """处理评论"""
        if session_id not in self.sessions:
            return
        
        # 广播评论
        await self.broadcast(session_id, {
            'type': 'comment',
            'comment': data.get('comment'),
            'author': participant_id
        })
    
    async def _handle_kick(self, participant_id: str, session_id: str,
                           data: dict) -> None:
        """处理踢出用户"""
        if session_id not in self.sessions:
            return
        
        session = self.sessions[session_id]
        
        # 只有主持人可以踢人
        if participant_id != session.host_id:
            return
        
        target_id = data.get('target_id')
        if target_id and target_id in session.participants:
            # 发送踢出通知
            if target_id in self.connections:
                await self.connections[target_id].send(json.dumps({
                    'type': 'kicked',
                    'message': '您已被主持人移出会话'
                }))
                await self.connections[target_id].close()
            
            del session.participants[target_id]
            await self._broadcast_presence(session_id)
    
    async def _broadcast_presence(self, session_id: str) -> None:
        """广播在线状态"""
        if session_id not in self.sessions:
            return
        
        session = self.sessions[session_id]
        
        await self.broadcast(session_id, {
            'type': 'presence_update',
            'participants': [p.to_dict() for p in session.participants.values()]
        })
    
    async def broadcast(self, session_id: str, message: dict, 
                        exclude: str = None) -> None:
        """广播消息给会话中的所有参与者
        
        Args:
            session_id: 会话 ID
            message: 消息内容
            exclude: 排除的参与者 ID
        """
        if session_id not in self.sessions:
            return
        
        session = self.sessions[session_id]
        message_str = json.dumps(message)
        
        for participant_id in session.participants:
            if participant_id == exclude:
                continue
            
            if participant_id in self.connections:
                try:
                    await self.connections[participant_id].send(message_str)
                except Exception:
                    pass
    
    def _generate_session_code(self) -> str:
        """生成 6 位会话码"""
        return secrets.token_hex(3).upper()
    
    def get_session(self, session_id: str) -> Optional[Session]:
        """获取会话信息"""
        return self.sessions.get(session_id)
    
    def close_session(self, session_id: str) -> None:
        """关闭会话"""
        if session_id in self.sessions:
            del self.sessions[session_id]
