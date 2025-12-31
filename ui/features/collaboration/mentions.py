# -*- coding: utf-8 -*-
"""@提及和任务管理模块"""

from dataclasses import dataclass
from typing import List, Optional, Callable
from datetime import datetime
import re
import uuid


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
    
    def to_dict(self) -> dict:
        return {
            'id': self.id,
            'author_id': self.author_id,
            'mentioned_id': self.mentioned_id,
            'position': self.position,
            'context': self.context,
            'created_at': self.created_at.isoformat(),
            'read': self.read
        }


@dataclass
class Task:
    """任务"""
    id: str
    description: str
    assignee_id: Optional[str]
    creator_id: str
    position: int
    completed: bool = False
    created_at: datetime = None
    completed_at: Optional[datetime] = None
    
    def __post_init__(self):
        if self.created_at is None:
            self.created_at = datetime.now()
    
    def to_dict(self) -> dict:
        return {
            'id': self.id,
            'description': self.description,
            'assignee_id': self.assignee_id,
            'creator_id': self.creator_id,
            'position': self.position,
            'completed': self.completed,
            'created_at': self.created_at.isoformat(),
            'completed_at': self.completed_at.isoformat() if self.completed_at else None
        }


class MentionManager:
    """@提及和任务管理器"""
    
    MENTION_PATTERN = re.compile(r'@(\w+)')
    TASK_PATTERN = re.compile(r'- \[([ xX])\] (.+?)(?:@(\w+))?$', re.MULTILINE)
    
    def __init__(self, app=None):
        self.app = app
        self.mentions: List[Mention] = []
        self.tasks: List[Task] = []
        self._notification_callback: Optional[Callable] = None
        self._participants: dict = {}  # name -> id mapping
    
    def set_participants(self, participants: dict) -> None:
        """设置参与者映射"""
        self._participants = participants
    
    def parse_mentions(self, content: str, author_id: str) -> List[Mention]:
        """解析文档中的@提及"""
        new_mentions = []
        
        for match in self.MENTION_PATTERN.finditer(content):
            mentioned_name = match.group(1)
            mentioned_id = self._participants.get(mentioned_name)
            
            if mentioned_id and mentioned_id != author_id:
                # 获取上下文
                start = max(0, match.start() - 30)
                end = min(len(content), match.end() + 30)
                context = content[start:end]
                
                mention = Mention(
                    id=str(uuid.uuid4()),
                    author_id=author_id,
                    mentioned_id=mentioned_id,
                    position=match.start(),
                    context=context,
                    created_at=datetime.now()
                )
                new_mentions.append(mention)
                self.mentions.append(mention)
        
        return new_mentions

    def parse_tasks(self, content: str, creator_id: str) -> List[Task]:
        """解析文档中的任务"""
        new_tasks = []
        
        for match in self.TASK_PATTERN.finditer(content):
            completed = match.group(1).lower() == 'x'
            description = match.group(2).strip()
            assignee_name = match.group(3)
            
            assignee_id = None
            if assignee_name:
                assignee_id = self._participants.get(assignee_name)
            
            task = Task(
                id=str(uuid.uuid4()),
                description=description,
                assignee_id=assignee_id,
                creator_id=creator_id,
                position=match.start(),
                completed=completed
            )
            new_tasks.append(task)
            self.tasks.append(task)
        
        return new_tasks
    
    def get_mentions_for_user(self, user_id: str) -> List[Mention]:
        """获取用户收到的提及"""
        return [m for m in self.mentions if m.mentioned_id == user_id]
    
    def get_unread_mentions(self, user_id: str) -> List[Mention]:
        """获取用户未读的提及"""
        return [m for m in self.mentions if m.mentioned_id == user_id and not m.read]
    
    def mark_mention_read(self, mention_id: str) -> None:
        """标记提及为已读"""
        for mention in self.mentions:
            if mention.id == mention_id:
                mention.read = True
                break
    
    def get_tasks_for_user(self, user_id: str) -> List[Task]:
        """获取分配给用户的任务"""
        return [t for t in self.tasks if t.assignee_id == user_id]
    
    def get_incomplete_tasks(self, user_id: str = None) -> List[Task]:
        """获取未完成的任务"""
        tasks = self.tasks
        if user_id:
            tasks = [t for t in tasks if t.assignee_id == user_id]
        return [t for t in tasks if not t.completed]
    
    def complete_task(self, task_id: str) -> Optional[Task]:
        """完成任务"""
        for task in self.tasks:
            if task.id == task_id:
                task.completed = True
                task.completed_at = datetime.now()
                
                # 通知创建者
                if self._notification_callback and task.creator_id:
                    self._notification_callback('task_completed', task)
                
                return task
        return None
    
    def uncomplete_task(self, task_id: str) -> Optional[Task]:
        """取消完成任务"""
        for task in self.tasks:
            if task.id == task_id:
                task.completed = False
                task.completed_at = None
                return task
        return None
    
    def notify_mention(self, mention: Mention) -> None:
        """发送提及通知"""
        if self._notification_callback:
            self._notification_callback('mention', mention)
    
    def on_notification(self, callback: Callable) -> None:
        """注册通知回调"""
        self._notification_callback = callback
    
    def get_user_suggestions(self, prefix: str) -> List[str]:
        """获取@自动完成建议"""
        prefix_lower = prefix.lower()
        suggestions = []
        
        for name in self._participants.keys():
            if name.lower().startswith(prefix_lower):
                suggestions.append(name)
        
        return suggestions[:10]
    
    def clear(self) -> None:
        """清空所有数据"""
        self.mentions = []
        self.tasks = []
