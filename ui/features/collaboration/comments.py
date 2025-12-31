# -*- coding: utf-8 -*-
"""评论和批注管理模块"""

from dataclasses import dataclass, field
from typing import List, Optional, Dict, Callable, Tuple
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
    
    def to_dict(self) -> dict:
        """转换为字典"""
        return {
            'id': self.id,
            'author_id': self.author_id,
            'author_name': self.author_name,
            'content': self.content,
            'created_at': self.created_at.isoformat(),
            'resolved': self.resolved,
            'replies': [r.to_dict() for r in self.replies]
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'Comment':
        """从字典创建"""
        return cls(
            id=data['id'],
            author_id=data['author_id'],
            author_name=data['author_name'],
            content=data['content'],
            created_at=datetime.fromisoformat(data['created_at']),
            resolved=data.get('resolved', False),
            replies=[cls.from_dict(r) for r in data.get('replies', [])]
        )


@dataclass
class CommentThread:
    """评论线程"""
    id: str
    document_range: Tuple[int, int]  # (start, end)
    comments: List[Comment] = field(default_factory=list)
    resolved: bool = False
    
    @property
    def comment_count(self) -> int:
        """评论数量（包括回复）"""
        count = len(self.comments)
        for comment in self.comments:
            count += len(comment.replies)
        return count
    
    def to_dict(self) -> dict:
        """转换为字典"""
        return {
            'id': self.id,
            'document_range': self.document_range,
            'comments': [c.to_dict() for c in self.comments],
            'resolved': self.resolved
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'CommentThread':
        """从字典创建"""
        return cls(
            id=data['id'],
            document_range=tuple(data['document_range']),
            comments=[Comment.from_dict(c) for c in data.get('comments', [])],
            resolved=data.get('resolved', False)
        )


class CommentManager:
    """评论和批注管理器"""
    
    def __init__(self, app=None):
        """初始化评论管理器
        
        Args:
            app: 应用实例
        """
        self.app = app
        self.threads: Dict[str, CommentThread] = {}
        self._on_comment_added: Optional[Callable] = None
        self._on_thread_resolved: Optional[Callable] = None
    
    def create_thread(self, start: int, end: int, initial_comment: str,
                      author_id: str, author_name: str) -> CommentThread:
        """创建评论线程
        
        Args:
            start: 文档范围起始位置
            end: 文档范围结束位置
            initial_comment: 初始评论内容
            author_id: 作者 ID
            author_name: 作者名称
            
        Returns:
            CommentThread 评论线程
        """
        thread_id = str(uuid.uuid4())
        
        comment = Comment(
            id=str(uuid.uuid4()),
            author_id=author_id,
            author_name=author_name,
            content=initial_comment,
            created_at=datetime.now()
        )
        
        thread = CommentThread(
            id=thread_id,
            document_range=(start, end),
            comments=[comment]
        )
        
        self.threads[thread_id] = thread
        
        if self._on_comment_added:
            self._on_comment_added(thread, comment)
        
        return thread
    
    def add_reply(self, thread_id: str, content: str,
                  author_id: str, author_name: str) -> Optional[Comment]:
        """添加回复
        
        Args:
            thread_id: 线程 ID
            content: 回复内容
            author_id: 作者 ID
            author_name: 作者名称
            
        Returns:
            Comment 评论对象
        """
        if thread_id not in self.threads:
            return None
        
        thread = self.threads[thread_id]
        
        comment = Comment(
            id=str(uuid.uuid4()),
            author_id=author_id,
            author_name=author_name,
            content=content,
            created_at=datetime.now()
        )
        
        thread.comments.append(comment)
        
        if self._on_comment_added:
            self._on_comment_added(thread, comment)
        
        return comment
    
    def resolve_thread(self, thread_id: str) -> None:
        """解决评论线程
        
        Args:
            thread_id: 线程 ID
        """
        if thread_id in self.threads:
            self.threads[thread_id].resolved = True
            
            if self._on_thread_resolved:
                self._on_thread_resolved(self.threads[thread_id])
    
    def unresolve_thread(self, thread_id: str) -> None:
        """重新打开评论线程
        
        Args:
            thread_id: 线程 ID
        """
        if thread_id in self.threads:
            self.threads[thread_id].resolved = False
    
    def delete_thread(self, thread_id: str) -> None:
        """删除评论线程
        
        Args:
            thread_id: 线程 ID
        """
        if thread_id in self.threads:
            del self.threads[thread_id]
    
    def delete_comment(self, thread_id: str, comment_id: str) -> None:
        """删除评论
        
        Args:
            thread_id: 线程 ID
            comment_id: 评论 ID
        """
        if thread_id not in self.threads:
            return
        
        thread = self.threads[thread_id]
        
        # 从主评论列表中删除
        thread.comments = [c for c in thread.comments if c.id != comment_id]
        
        # 从回复中删除
        for comment in thread.comments:
            comment.replies = [r for r in comment.replies if r.id != comment_id]
        
        # 如果线程没有评论了，删除线程
        if not thread.comments:
            del self.threads[thread_id]
    
    def get_thread(self, thread_id: str) -> Optional[CommentThread]:
        """获取评论线程
        
        Args:
            thread_id: 线程 ID
            
        Returns:
            CommentThread 或 None
        """
        return self.threads.get(thread_id)
    
    def get_threads_in_range(self, start: int, end: int) -> List[CommentThread]:
        """获取指定范围内的评论线程
        
        Args:
            start: 范围起始位置
            end: 范围结束位置
            
        Returns:
            CommentThread 列表
        """
        result = []
        
        for thread in self.threads.values():
            t_start, t_end = thread.document_range
            
            # 检查范围是否重叠
            if t_start <= end and t_end >= start:
                result.append(thread)
        
        return result
    
    def get_all_threads(self, include_resolved: bool = True) -> List[CommentThread]:
        """获取所有评论线程
        
        Args:
            include_resolved: 是否包含已解决的线程
            
        Returns:
            CommentThread 列表
        """
        if include_resolved:
            return list(self.threads.values())
        return [t for t in self.threads.values() if not t.resolved]
    
    def export_comments(self, include_resolved: bool = False) -> str:
        """导出评论为 Markdown
        
        Args:
            include_resolved: 是否包含已解决的评论
            
        Returns:
            Markdown 格式的评论
        """
        lines = ["# 评论", ""]
        
        for thread in self.get_all_threads(include_resolved):
            status = "✅ 已解决" if thread.resolved else "💬 进行中"
            lines.append(f"## {status} (位置: {thread.document_range[0]}-{thread.document_range[1]})")
            lines.append("")
            
            for comment in thread.comments:
                lines.append(f"**{comment.author_name}** ({comment.created_at.strftime('%Y-%m-%d %H:%M')}):")
                lines.append(f"> {comment.content}")
                lines.append("")
                
                for reply in comment.replies:
                    lines.append(f"  - **{reply.author_name}**: {reply.content}")
            
            lines.append("")
        
        return "\n".join(lines)
    
    def sync_with_server(self, server_threads: List[dict]) -> None:
        """与服务器同步评论
        
        Args:
            server_threads: 服务器端的评论线程数据
        """
        for thread_data in server_threads:
            thread = CommentThread.from_dict(thread_data)
            self.threads[thread.id] = thread
    
    def on_comment_added(self, callback: Callable) -> None:
        """注册评论添加回调"""
        self._on_comment_added = callback
    
    def on_thread_resolved(self, callback: Callable) -> None:
        """注册线程解决回调"""
        self._on_thread_resolved = callback
    
    def update_range_on_edit(self, position: int, delta: int) -> None:
        """编辑时更新评论范围
        
        Args:
            position: 编辑位置
            delta: 变化量（正数为插入，负数为删除）
        """
        for thread in self.threads.values():
            start, end = thread.document_range
            
            if position <= start:
                # 编辑在评论范围之前
                thread.document_range = (start + delta, end + delta)
            elif position < end:
                # 编辑在评论范围内
                thread.document_range = (start, max(start, end + delta))
