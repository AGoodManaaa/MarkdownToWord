# Design Document: 实时协作编辑

## Overview

本设计文档描述了实时协作编辑功能的技术实现方案。该功能支持局域网内多人同时编辑文档，使用 WebSocket 进行实时通信，CRDT 算法解决并发冲突。

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                    Collaboration System                          │
├─────────────────────────────────────────────────────────────────┤
│                                                                  │
│  ┌─────────────┐     ┌─────────────┐     ┌─────────────┐       │
│  │   Client    │────▶│  WebSocket  │◀────│   Client    │       │
│  │  (Editor)   │     │   Server    │     │  (Editor)   │       │
│  └─────────────┘     └─────────────┘     └─────────────┘       │
│        │                   │                   │                │
│        ▼                   ▼                   ▼                │
│  ┌─────────────┐     ┌─────────────┐     ┌─────────────┐       │
│  │    CRDT     │     │   Session   │     │    CRDT     │       │
│  │   Engine    │     │   Manager   │     │   Engine    │       │
│  └─────────────┘     └─────────────┘     └─────────────┘       │
│                            │                                    │
│                            ▼                                    │
│                     ┌─────────────┐                             │
│                     │  Comments   │                             │
│                     │  & History  │                             │
│                     └─────────────┘                             │
└─────────────────────────────────────────────────────────────────┘
```

## Components and Interfaces

### 1. CollaborationServer - 协作服务器

```python
import asyncio
import websockets
from typing import Dict, Set, Optional
from dataclasses import dataclass
import json
import secrets

@dataclass
class Session:
    """协作会话"""
    id: str
    host_id: str
    document_content: str
    participants: Set[str]
    password_hash: Optional[str]
    created_at: float
    crdt_state: bytes

@dataclass
class Participant:
    """参与者信息"""
    id: str
    name: str
    color: str
    cursor_position: int
    selection: Optional[tuple]
    permission: str  # 'edit', 'comment', 'view'
    is_online: bool
    last_active: float

class CollaborationServer:
    """协作服务器（WebSocket）"""
    
    def __init__(self, host: str = "0.0.0.0", port: int = 8765):
        self.host = host
        self.port = port
        self.sessions: Dict[str, Session] = {}
        self.connections: Dict[str, websockets.WebSocketServerProtocol] = {}
        self._server = None
    
    async def start(self) -> str:
        """启动服务器，返回连接地址"""
        pass
    
    async def stop(self) -> None:
        """停止服务器"""
        pass
    
    def create_session(self, document: str, password: Optional[str] = None) -> str:
        """创建新会话，返回会话码"""
        pass
    
    async def handle_connection(self, websocket, path) -> None:
        """处理 WebSocket 连接"""
        pass
    
    async def broadcast(self, session_id: str, message: dict, exclude: str = None) -> None:
        """广播消息给会话中的所有参与者"""
        pass
    
    async def handle_message(self, participant_id: str, message: dict) -> None:
        """处理收到的消息"""
        pass
    
    def _generate_session_code(self) -> str:
        """生成 6 位会话码"""
        return secrets.token_hex(3).upper()
```

### 2. CollaborationClient - 协作客户端

```python
import asyncio
import websockets
from typing import Callable, Optional

class CollaborationClient:
    """协作客户端"""
    
    def __init__(self, app):
        self.app = app
        self.websocket = None
        self.session_id: Optional[str] = None
        self.participant_id: Optional[str] = None
        self.crdt_engine = None
        self._message_handlers: Dict[str, Callable] = {}
    
    async def connect(self, address: str, session_code: str, 
                      name: str, password: Optional[str] = None) -> bool:
        """连接到协作会话"""
        pass
    
    async def disconnect(self) -> None:
        """断开连接"""
        pass
    
    async def send_operation(self, operation: dict) -> None:
        """发送编辑操作"""
        pass
    
    async def send_cursor_update(self, position: int, selection: Optional[tuple] = None) -> None:
        """发送光标位置更新"""
        pass
    
    def on_remote_operation(self, callback: Callable[[dict], None]) -> None:
        """注册远程操作回调"""
        pass
    
    def on_cursor_update(self, callback: Callable[[str, int, Optional[tuple]], None]) -> None:
        """注册光标更新回调"""
        pass
    
    def on_participant_change(self, callback: Callable[[str, str], None]) -> None:
        """注册参与者变化回调 (participant_id, event)"""
        pass
    
    async def _receive_loop(self) -> None:
        """消息接收循环"""
        pass
```

### 3. CRDTEngine - CRDT 引擎

```python
from typing import List, Tuple, Optional
from dataclasses import dataclass
import uuid

@dataclass
class CRDTOperation:
    """CRDT 操作"""
    id: str
    type: str  # 'insert', 'delete'
    position: int
    content: str  # for insert
    length: int   # for delete
    author: str
    timestamp: float
    vector_clock: dict

class CRDTEngine:
    """
    CRDT 引擎 - 使用 RGA (Replicated Growable Array) 算法
    实现无冲突的并发文本编辑
    """
    
    def __init__(self, site_id: str):
        self.site_id = site_id
        self._document: List[tuple] = []  # [(char, id, tombstone), ...]
        self._vector_clock: Dict[str, int] = {}
        self._pending_ops: List[CRDTOperation] = []
    
    def local_insert(self, position: int, content: str) -> CRDTOperation:
        """本地插入操作"""
        pass
    
    def local_delete(self, position: int, length: int) -> CRDTOperation:
        """本地删除操作"""
        pass
    
    def apply_remote(self, operation: CRDTOperation) -> Tuple[int, str]:
        """应用远程操作，返回 (position, content)"""
        pass
    
    def get_content(self) -> str:
        """获取当前文档内容"""
        pass
    
    def get_state(self) -> bytes:
        """序列化当前状态"""
        pass
    
    def load_state(self, state: bytes) -> None:
        """加载状态"""
        pass
    
    def _find_insert_position(self, op: CRDTOperation) -> int:
        """找到插入位置（处理并发）"""
        pass
    
    def _increment_clock(self) -> dict:
        """递增向量时钟"""
        pass
```

### 4. CursorManager - 光标管理器

```python
from typing import Dict, Optional, Tuple
import colorsys

@dataclass
class RemoteCursor:
    """远程光标"""
    participant_id: str
    name: str
    color: str
    position: int
    selection_start: Optional[int]
    selection_end: Optional[int]
    last_update: float

class CursorManager:
    """远程光标管理器"""
    
    CURSOR_COLORS = [
        "#FF6B6B", "#4ECDC4", "#45B7D1", "#96CEB4",
        "#FFEAA7", "#DDA0DD", "#98D8C8", "#F7DC6F"
    ]
    
    def __init__(self, editor_widget):
        self.editor = editor_widget
        self.cursors: Dict[str, RemoteCursor] = {}
        self._color_index = 0
    
    def add_cursor(self, participant_id: str, name: str) -> str:
        """添加远程光标，返回分配的颜色"""
        pass
    
    def remove_cursor(self, participant_id: str) -> None:
        """移除远程光标"""
        pass
    
    def update_cursor(self, participant_id: str, position: int,
                      selection: Optional[Tuple[int, int]] = None) -> None:
        """更新光标位置"""
        pass
    
    def render_cursors(self) -> None:
        """渲染所有远程光标"""
        pass
    
    def _draw_cursor(self, cursor: RemoteCursor) -> None:
        """绘制单个光标"""
        pass
    
    def _draw_selection(self, cursor: RemoteCursor) -> None:
        """绘制选区高亮"""
        pass
    
    def _get_next_color(self) -> str:
        """获取下一个可用颜色"""
        pass
```

### 5. CommentManager - 评论管理器

```python
from dataclasses import dataclass, field
from typing import List, Optional
from datetime import datetime
import uuid

@dataclass
class Comment:
    """评论"""
    id: str
    author_id: str
    author_name: str
    content: str
    created_at: datetime
    resolved: bool = False
    replies: List['Comment'] = field(default_factory=list)

@dataclass
class CommentThread:
    """评论线程"""
    id: str
    document_range: Tuple[int, int]  # (start, end)
    comments: List[Comment]
    resolved: bool = False

class CommentManager:
    """评论和批注管理器"""
    
    def __init__(self, app):
        self.app = app
        self.threads: Dict[str, CommentThread] = {}
        self._on_comment_added: Optional[Callable] = None
    
    def create_thread(self, start: int, end: int, initial_comment: str,
                      author_id: str, author_name: str) -> CommentThread:
        """创建评论线程"""
        pass
    
    def add_reply(self, thread_id: str, content: str,
                  author_id: str, author_name: str) -> Comment:
        """添加回复"""
        pass
    
    def resolve_thread(self, thread_id: str) -> None:
        """解决评论线程"""
        pass
    
    def delete_comment(self, thread_id: str, comment_id: str) -> None:
        """删除评论"""
        pass
    
    def get_threads_in_range(self, start: int, end: int) -> List[CommentThread]:
        """获取指定范围内的评论线程"""
        pass
    
    def export_comments(self, include_resolved: bool = False) -> str:
        """导出评论为 Markdown"""
        pass
    
    def sync_with_server(self, server_threads: List[dict]) -> None:
        """与服务器同步评论"""
        pass
```

### 6. HistoryManager - 历史管理器

```python
from dataclasses import dataclass
from typing import List, Optional
from datetime import datetime

@dataclass
class HistoryEntry:
    """历史条目"""
    id: str
    timestamp: datetime
    author_id: str
    author_name: str
    operation_type: str  # 'edit', 'comment', 'resolve'
    content_before: str
    content_after: str
    range: Tuple[int, int]

class HistoryManager:
    """修改历史管理器"""
    
    def __init__(self, max_entries: int = 1000):
        self.max_entries = max_entries
        self.entries: List[HistoryEntry] = []
        self._snapshots: Dict[str, str] = {}  # 定期快照
    
    def record(self, entry: HistoryEntry) -> None:
        """记录历史条目"""
        pass
    
    def get_history(self, limit: int = 50) -> List[HistoryEntry]:
        """获取历史记录"""
        pass
    
    def get_entry(self, entry_id: str) -> Optional[HistoryEntry]:
        """获取指定条目"""
        pass
    
    def restore(self, entry_id: str) -> str:
        """恢复到指定版本，返回恢复后的内容"""
        pass
    
    def diff(self, entry_id1: str, entry_id2: str) -> List[dict]:
        """对比两个版本的差异"""
        pass
    
    def create_snapshot(self, content: str) -> str:
        """创建快照，返回快照 ID"""
        pass
    
    def compress_old_entries(self) -> None:
        """压缩旧条目"""
        pass
    
    def export_history(self) -> bytes:
        """导出历史（用于保存）"""
        pass
    
    def import_history(self, data: bytes) -> None:
        """导入历史"""
        pass
```

### 7. MentionManager - @提及管理器

```python
@dataclass
class Mention:
    """@提及"""
    id: str
    author_id: str
    mentioned_id: str
    position: int
    context: str
    created_at: datetime
    read: bool = False

@dataclass
class Task:
    """任务"""
    id: str
    description: str
    assignee_id: Optional[str]
    creator_id: str
    position: int
    completed: bool = False
    created_at: datetime
    completed_at: Optional[datetime] = None

class MentionManager:
    """@提及和任务管理器"""
    
    MENTION_PATTERN = re.compile(r'@(\w+)')
    TASK_PATTERN = re.compile(r'- \[([ x])\] (.+?)(?:@(\w+))?$', re.MULTILINE)
    
    def __init__(self, app):
        self.app = app
        self.mentions: List[Mention] = []
        self.tasks: List[Task] = []
        self._notification_callback: Optional[Callable] = None
    
    def parse_mentions(self, content: str, author_id: str) -> List[Mention]:
        """解析文档中的@提及"""
        pass
    
    def parse_tasks(self, content: str, creator_id: str) -> List[Task]:
        """解析文档中的任务"""
        pass
    
    def get_mentions_for_user(self, user_id: str) -> List[Mention]:
        """获取用户收到的提及"""
        pass
    
    def get_tasks_for_user(self, user_id: str) -> List[Task]:
        """获取分配给用户的任务"""
        pass
    
    def complete_task(self, task_id: str) -> None:
        """完成任务"""
        pass
    
    def notify_mention(self, mention: Mention) -> None:
        """发送提及通知"""
        pass
    
    def get_user_suggestions(self, prefix: str) -> List[Participant]:
        """获取@自动完成建议"""
        pass
```

## Data Models

### 消息协议

```python
# WebSocket 消息格式
message_types = {
    # 连接管理
    'join': {'session_code': str, 'name': str, 'password': str},
    'leave': {},
    'kick': {'participant_id': str},
    
    # 文档同步
    'sync_request': {},
    'sync_response': {'content': str, 'crdt_state': bytes},
    'operation': {'op': CRDTOperation},
    
    # 光标同步
    'cursor_update': {'position': int, 'selection': tuple},
    
    # 在线状态
    'presence_update': {'participants': List[Participant]},
    
    # 评论
    'comment_add': {'thread_id': str, 'comment': Comment},
    'comment_resolve': {'thread_id': str},
    
    # 提及
    'mention': {'mention': Mention},
    'task_update': {'task': Task}
}
```

### 配置结构

```python
collaboration_config = {
    'server': {
        'host': '0.0.0.0',
        'port': 8765,
        'max_participants': 10,
        'session_timeout_minutes': 60
    },
    'sync': {
        'debounce_ms': 50,
        'batch_operations': True,
        'max_batch_size': 100
    },
    'cursor': {
        'update_interval_ms': 100,
        'show_names': True,
        'fade_inactive_ms': 5000
    },
    'history': {
        'max_entries': 1000,
        'snapshot_interval': 100,
        'compress_after_days': 7
    },
    'security': {
        'require_password': False,
        'encrypt_traffic': True,
        'max_message_size_kb': 1024
    }
}
```

## Correctness Properties

*A property is a characteristic or behavior that should hold true across all valid executions of a system-essentially, a formal statement about what the system should do. Properties serve as the bridge between human-readable specifications and machine-verifiable correctness guarantees.*

### Property 1: CRDT 收敛性

*For any* 一系列并发操作，所有客户端最终应收敛到相同的文档状态
**Validates: Requirements 2.5**

### Property 2: 操作顺序保持

*For any* 单个客户端的操作序列，应用到其他客户端后应保持相对顺序
**Validates: Requirements 2.3**

### Property 3: 光标位置一致性

*For any* 远程操作应用后，本地光标位置应正确调整以保持相对位置
**Validates: Requirements 2.4**

### Property 4: 评论范围有效性

*For any* 评论线程，其关联的文档范围应始终有效（在文档边界内）
**Validates: Requirements 4.2**

### Property 5: 历史可恢复性

*For any* 历史条目，恢复操作应产生与该时间点完全相同的文档状态
**Validates: Requirements 6.5**

### Property 6: 会话隔离性

*For any* 两个不同的会话，它们的操作不应相互影响
**Validates: Requirements 1.1**

## Error Handling

```python
class CollaborationError(Exception):
    """协作相关错误基类"""
    pass

class ConnectionError(CollaborationError):
    """连接错误"""
    pass

class SyncError(CollaborationError):
    """同步错误"""
    pass

class PermissionError(CollaborationError):
    """权限错误"""
    pass

class SessionError(CollaborationError):
    """会话错误"""
    pass
```

## Testing Strategy

### 测试覆盖

| 组件 | 单元测试 | 属性测试 |
| ---- | -------- | -------- |
| CollaborationServer | ✓ | ✓ (Property 6) |
| CRDTEngine | ✓ | ✓ (Property 1, 2) |
| CursorManager | ✓ | ✓ (Property 3) |
| CommentManager | ✓ | ✓ (Property 4) |
| HistoryManager | ✓ | ✓ (Property 5) |

## Dependencies

```
websockets>=11.0
asyncio
```
