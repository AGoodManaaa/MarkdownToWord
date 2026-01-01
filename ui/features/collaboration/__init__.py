# -*- coding: utf-8 -*-
"""实时协作编辑功能模块 - 腾讯文档风格增强版"""

from .crdt import CRDTEngine, CRDTOperation
from .server import CollaborationServer, Session, Participant
from .client import CollaborationClient
from .cursor import CursorManager, RemoteCursor, CursorLabel, CursorLine
from .comments import CommentManager, Comment, CommentThread
from .history import HistoryManager, HistoryEntry
from .mentions import MentionManager, Mention, Task
from .panels import CollaborationFeature
from .comment_panel import CommentSidePanel, CommentBubble, CommentData, REACTIONS
from .notifications import NotificationManager, NotificationPanel, NotificationType, Notification

__all__ = [
    'CRDTEngine',
    'CRDTOperation',
    'CollaborationServer',
    'Session',
    'Participant',
    'CollaborationClient',
    'CursorManager',
    'RemoteCursor',
    'CursorLabel',
    'CursorLine',
    'CommentManager',
    'Comment',
    'CommentThread',
    'CommentSidePanel',
    'CommentBubble',
    'CommentData',
    'REACTIONS',
    'HistoryManager',
    'HistoryEntry',
    'MentionManager',
    'Mention',
    'Task',
    'CollaborationFeature',
    'NotificationManager',
    'NotificationPanel',
    'NotificationType',
    'Notification',
]

