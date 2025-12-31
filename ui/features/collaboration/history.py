# -*- coding: utf-8 -*-
"""修改历史管理模块"""

from dataclasses import dataclass
from typing import List, Optional, Dict, Tuple
from datetime import datetime
import uuid
import difflib
import json
import gzip


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
    range: Tuple[int, int] = (0, 0)
    
    def to_dict(self) -> dict:
        """转换为字典"""
        return {
            'id': self.id,
            'timestamp': self.timestamp.isoformat(),
            'author_id': self.author_id,
            'author_name': self.author_name,
            'operation_type': self.operation_type,
            'content_before': self.content_before,
            'content_after': self.content_after,
            'range': self.range
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'HistoryEntry':
        """从字典创建"""
        return cls(
            id=data['id'],
            timestamp=datetime.fromisoformat(data['timestamp']),
            author_id=data['author_id'],
            author_name=data['author_name'],
            operation_type=data['operation_type'],
            content_before=data['content_before'],
            content_after=data['content_after'],
            range=tuple(data.get('range', (0, 0)))
        )


class HistoryManager:
    """修改历史管理器"""
    
    def __init__(self, max_entries: int = 1000):
        """初始化历史管理器
        
        Args:
            max_entries: 最大历史条目数
        """
        self.max_entries = max_entries
        self.entries: List[HistoryEntry] = []
        self._snapshots: Dict[str, str] = {}  # snapshot_id -> content
        self._snapshot_interval = 50  # 每 50 个条目创建一个快照
    
    def record(self, author_id: str, author_name: str,
               operation_type: str, content_before: str,
               content_after: str, range: Tuple[int, int] = (0, 0)) -> HistoryEntry:
        """记录历史条目
        
        Args:
            author_id: 作者 ID
            author_name: 作者名称
            operation_type: 操作类型
            content_before: 修改前内容
            content_after: 修改后内容
            range: 修改范围
            
        Returns:
            HistoryEntry 历史条目
        """
        entry = HistoryEntry(
            id=str(uuid.uuid4()),
            timestamp=datetime.now(),
            author_id=author_id,
            author_name=author_name,
            operation_type=operation_type,
            content_before=content_before,
            content_after=content_after,
            range=range
        )
        
        self.entries.append(entry)
        
        # 检查是否需要创建快照
        if len(self.entries) % self._snapshot_interval == 0:
            self.create_snapshot(content_after)
        
        # 检查是否超过最大条目数
        if len(self.entries) > self.max_entries:
            self.compress_old_entries()
        
        return entry

    def get_history(self, limit: int = 50) -> List[HistoryEntry]:
        """获取历史记录
        
        Args:
            limit: 返回数量限制
            
        Returns:
            HistoryEntry 列表（最新的在前）
        """
        return list(reversed(self.entries[-limit:]))
    
    def get_entry(self, entry_id: str) -> Optional[HistoryEntry]:
        """获取指定条目
        
        Args:
            entry_id: 条目 ID
            
        Returns:
            HistoryEntry 或 None
        """
        for entry in self.entries:
            if entry.id == entry_id:
                return entry
        return None
    
    def restore(self, entry_id: str) -> Optional[str]:
        """恢复到指定版本
        
        Args:
            entry_id: 条目 ID
            
        Returns:
            恢复后的内容
        """
        entry = self.get_entry(entry_id)
        if entry:
            return entry.content_after
        return None
    
    def diff(self, entry_id1: str, entry_id2: str) -> List[dict]:
        """对比两个版本的差异
        
        Args:
            entry_id1: 第一个条目 ID
            entry_id2: 第二个条目 ID
            
        Returns:
            差异列表
        """
        entry1 = self.get_entry(entry_id1)
        entry2 = self.get_entry(entry_id2)
        
        if not entry1 or not entry2:
            return []
        
        content1 = entry1.content_after
        content2 = entry2.content_after
        
        differ = difflib.unified_diff(
            content1.splitlines(keepends=True),
            content2.splitlines(keepends=True),
            lineterm=''
        )
        
        result = []
        for line in differ:
            if line.startswith('+') and not line.startswith('+++'):
                result.append({'type': 'add', 'content': line[1:]})
            elif line.startswith('-') and not line.startswith('---'):
                result.append({'type': 'remove', 'content': line[1:]})
            elif not line.startswith(('@@', '---', '+++')):
                result.append({'type': 'unchanged', 'content': line})
        
        return result

    def create_snapshot(self, content: str) -> str:
        """创建快照
        
        Args:
            content: 文档内容
            
        Returns:
            快照 ID
        """
        snapshot_id = str(uuid.uuid4())
        self._snapshots[snapshot_id] = content
        return snapshot_id
    
    def compress_old_entries(self) -> None:
        """压缩旧条目"""
        if len(self.entries) <= self.max_entries:
            return
        
        # 保留最近的条目
        keep_count = self.max_entries // 2
        self.entries = self.entries[-keep_count:]
    
    def export_history(self) -> bytes:
        """导出历史
        
        Returns:
            压缩的历史数据
        """
        data = {
            'entries': [e.to_dict() for e in self.entries],
            'snapshots': self._snapshots
        }
        json_str = json.dumps(data, ensure_ascii=False)
        return gzip.compress(json_str.encode('utf-8'))
    
    def import_history(self, data: bytes) -> None:
        """导入历史
        
        Args:
            data: 压缩的历史数据
        """
        json_str = gzip.decompress(data).decode('utf-8')
        loaded = json.loads(json_str)
        
        self.entries = [HistoryEntry.from_dict(e) for e in loaded.get('entries', [])]
        self._snapshots = loaded.get('snapshots', {})
    
    def clear(self) -> None:
        """清空历史"""
        self.entries = []
        self._snapshots = {}
    
    def get_entries_by_author(self, author_id: str) -> List[HistoryEntry]:
        """获取指定作者的历史条目"""
        return [e for e in self.entries if e.author_id == author_id]
    
    def get_entries_in_range(self, start: datetime, end: datetime) -> List[HistoryEntry]:
        """获取指定时间范围内的历史条目"""
        return [e for e in self.entries if start <= e.timestamp <= end]
