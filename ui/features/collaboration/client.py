# -*- coding: utf-8 -*-
"""协作客户端模块"""

import asyncio
import json
import threading
from typing import Callable, Optional, Dict, Any

try:
    import websockets
    WEBSOCKETS_AVAILABLE = True
except ImportError:
    WEBSOCKETS_AVAILABLE = False
    websockets = None

from .crdt import CRDTEngine, CRDTOperation


class CollaborationClient:
    """协作客户端"""
    
    def __init__(self, app=None):
        """初始化协作客户端
        
        Args:
            app: 应用实例
        """
        self.app = app
        self.websocket = None
        self.session_id: Optional[str] = None
        self.participant_id: Optional[str] = None
        self.crdt_engine: Optional[CRDTEngine] = None
        
        self._message_handlers: Dict[str, Callable] = {}
        self._connected = False
        self._receive_task = None
        self._loop = None
        self._thread = None
    
    @property
    def is_connected(self) -> bool:
        """是否已连接"""
        return self._connected and self.websocket is not None
    
    def connect(self, address: str, session_code: str,
                name: str, password: Optional[str] = None,
                on_success: Callable = None,
                on_error: Callable = None) -> None:
        """连接到协作会话（异步）
        
        Args:
            address: 服务器地址
            session_code: 会话码
            name: 用户名
            password: 密码（可选）
            on_success: 成功回调
            on_error: 错误回调
        """
        if not WEBSOCKETS_AVAILABLE:
            if on_error:
                on_error("websockets 库未安装")
            return
        
        def run_async():
            self._loop = asyncio.new_event_loop()
            asyncio.set_event_loop(self._loop)
            
            try:
                self._loop.run_until_complete(
                    self._connect_async(address, session_code, name, password)
                )
                
                if self._connected and on_success:
                    self.app.after(0, on_success) if self.app else on_success()
                
                # 启动接收循环
                self._loop.run_until_complete(self._receive_loop())
                
            except Exception as e:
                if on_error:
                    error_msg = str(e)
                    if self.app:
                        self.app.after(0, lambda: on_error(error_msg))
                    else:
                        on_error(error_msg)
            finally:
                self._loop.close()
        
        self._thread = threading.Thread(target=run_async, daemon=True)
        self._thread.start()
    
    async def _connect_async(self, address: str, session_code: str,
                             name: str, password: Optional[str]) -> bool:
        """异步连接"""
        try:
            self.websocket = await websockets.connect(address)
            
            # 发送加入请求
            await self.websocket.send(json.dumps({
                'type': 'join',
                'session_code': session_code,
                'name': name,
                'password': password
            }))
            
            # 等待响应
            response = await self.websocket.recv()
            data = json.loads(response)
            
            if data.get('type') == 'join_success':
                self.session_id = data.get('session_id')
                self.participant_id = data.get('participant_id')
                self._connected = True
                
                # 初始化 CRDT 引擎
                self.crdt_engine = CRDTEngine(self.participant_id)
                self.crdt_engine.set_content(data.get('document', ''))
                
                # 触发初始化回调
                if 'init' in self._message_handlers:
                    self._message_handlers['init'](data)
                
                return True
            else:
                error_msg = data.get('message', '连接失败')
                raise Exception(error_msg)
                
        except Exception as e:
            self._connected = False
            raise
    
    def disconnect(self) -> None:
        """断开连接"""
        self._connected = False
        
        if self._loop and self.websocket:
            asyncio.run_coroutine_threadsafe(
                self._disconnect_async(),
                self._loop
            )
    
    async def _disconnect_async(self) -> None:
        """异步断开连接"""
        if self.websocket:
            try:
                await self.websocket.send(json.dumps({'type': 'leave'}))
                await self.websocket.close()
            except Exception:
                pass
            finally:
                self.websocket = None
    
    def send_operation(self, operation: CRDTOperation) -> None:
        """发送编辑操作
        
        Args:
            operation: CRDT 操作
        """
        if not self.is_connected:
            return
        
        asyncio.run_coroutine_threadsafe(
            self._send_async({
                'type': 'operation',
                'operation': operation.to_dict()
            }),
            self._loop
        )
    
    def send_cursor_update(self, position: int, 
                           selection: Optional[tuple] = None) -> None:
        """发送光标位置更新
        
        Args:
            position: 光标位置
            selection: 选区 (start, end)
        """
        if not self.is_connected:
            return
        
        asyncio.run_coroutine_threadsafe(
            self._send_async({
                'type': 'cursor_update',
                'position': position,
                'selection': selection
            }),
            self._loop
        )
    
    def send_comment(self, thread_id: str, content: str) -> None:
        """发送评论
        
        Args:
            thread_id: 评论线程 ID
            content: 评论内容
        """
        if not self.is_connected:
            return
        
        asyncio.run_coroutine_threadsafe(
            self._send_async({
                'type': 'comment',
                'comment': {
                    'thread_id': thread_id,
                    'content': content
                }
            }),
            self._loop
        )
    
    async def _send_async(self, message: dict) -> None:
        """异步发送消息"""
        if self.websocket:
            try:
                await self.websocket.send(json.dumps(message))
            except Exception:
                pass
    
    def on_remote_operation(self, callback: Callable[[dict], None]) -> None:
        """注册远程操作回调
        
        Args:
            callback: 回调函数 (operation_dict)
        """
        self._message_handlers['operation'] = callback
    
    def on_cursor_update(self, callback: Callable[[str, int, Optional[tuple]], None]) -> None:
        """注册光标更新回调
        
        Args:
            callback: 回调函数 (participant_id, position, selection)
        """
        self._message_handlers['cursor_update'] = callback
    
    def on_participant_change(self, callback: Callable[[list], None]) -> None:
        """注册参与者变化回调
        
        Args:
            callback: 回调函数 (participants_list)
        """
        self._message_handlers['presence_update'] = callback
    
    def on_comment(self, callback: Callable[[dict], None]) -> None:
        """注册评论回调
        
        Args:
            callback: 回调函数 (comment_dict)
        """
        self._message_handlers['comment'] = callback
    
    def on_init(self, callback: Callable[[dict], None]) -> None:
        """注册初始化回调
        
        Args:
            callback: 回调函数 (init_data)
        """
        self._message_handlers['init'] = callback
    
    def on_kicked(self, callback: Callable[[str], None]) -> None:
        """注册被踢出回调
        
        Args:
            callback: 回调函数 (message)
        """
        self._message_handlers['kicked'] = callback
    
    async def _receive_loop(self) -> None:
        """消息接收循环"""
        if not self.websocket:
            return
        
        try:
            async for message in self.websocket:
                if not self._connected:
                    break
                
                try:
                    data = json.loads(message)
                    msg_type = data.get('type')
                    
                    # 调用对应的处理器
                    if msg_type in self._message_handlers:
                        handler = self._message_handlers[msg_type]
                        
                        if msg_type == 'operation':
                            # 应用远程操作到 CRDT
                            op_data = data.get('operation', {})
                            if self.crdt_engine:
                                op = CRDTOperation.from_dict(op_data)
                                self.crdt_engine.apply_remote(op)
                            
                            if self.app:
                                self.app.after(0, lambda d=data: handler(d.get('operation')))
                            else:
                                handler(data.get('operation'))
                        
                        elif msg_type == 'cursor_update':
                            if self.app:
                                self.app.after(0, lambda d=data: handler(
                                    d.get('participant_id'),
                                    d.get('position', 0),
                                    d.get('selection')
                                ))
                            else:
                                handler(
                                    data.get('participant_id'),
                                    data.get('position', 0),
                                    data.get('selection')
                                )
                        
                        elif msg_type == 'presence_update':
                            if self.app:
                                self.app.after(0, lambda d=data: handler(d.get('participants', [])))
                            else:
                                handler(data.get('participants', []))
                        
                        elif msg_type == 'comment':
                            if self.app:
                                self.app.after(0, lambda d=data: handler(d.get('comment')))
                            else:
                                handler(data.get('comment'))
                        
                        elif msg_type == 'kicked':
                            self._connected = False
                            if self.app:
                                self.app.after(0, lambda d=data: handler(d.get('message', '')))
                            else:
                                handler(data.get('message', ''))
                            break
                
                except json.JSONDecodeError:
                    pass
                except Exception:
                    pass
        
        except websockets.exceptions.ConnectionClosed:
            self._connected = False
        except Exception:
            self._connected = False
    
    def local_insert(self, position: int, content: str) -> Optional[CRDTOperation]:
        """本地插入操作
        
        Args:
            position: 插入位置
            content: 插入内容
            
        Returns:
            CRDT 操作
        """
        if self.crdt_engine:
            op = self.crdt_engine.local_insert(position, content)
            self.send_operation(op)
            return op
        return None
    
    def local_delete(self, position: int, length: int) -> Optional[CRDTOperation]:
        """本地删除操作
        
        Args:
            position: 删除位置
            length: 删除长度
            
        Returns:
            CRDT 操作
        """
        if self.crdt_engine:
            op = self.crdt_engine.local_delete(position, length)
            self.send_operation(op)
            return op
        return None
    
    def get_document(self) -> str:
        """获取当前文档内容"""
        if self.crdt_engine:
            return self.crdt_engine.get_content()
        return ""
