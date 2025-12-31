# -*- coding: utf-8 -*-
"""CRDT 引擎模块 - 实现无冲突并发编辑"""

from dataclasses import dataclass, field
from typing import List, Tuple, Optional, Dict, Any
import uuid
import time
import json
import pickle


@dataclass
class CRDTOperation:
    """CRDT 操作"""
    id: str
    type: str  # 'insert', 'delete'
    position: int
    content: str = ""  # for insert
    length: int = 0    # for delete
    author: str = ""
    timestamp: float = 0.0
    vector_clock: Dict[str, int] = field(default_factory=dict)
    
    def to_dict(self) -> dict:
        """转换为字典"""
        return {
            'id': self.id,
            'type': self.type,
            'position': self.position,
            'content': self.content,
            'length': self.length,
            'author': self.author,
            'timestamp': self.timestamp,
            'vector_clock': self.vector_clock
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'CRDTOperation':
        """从字典创建"""
        return cls(
            id=data['id'],
            type=data['type'],
            position=data['position'],
            content=data.get('content', ''),
            length=data.get('length', 0),
            author=data.get('author', ''),
            timestamp=data.get('timestamp', 0.0),
            vector_clock=data.get('vector_clock', {})
        )


@dataclass
class CRDTChar:
    """CRDT 字符（RGA 算法）"""
    char: str
    id: Tuple[str, int]  # (site_id, sequence)
    visible: bool = True
    
    def __lt__(self, other: 'CRDTChar') -> bool:
        """比较顺序（用于排序）"""
        # 先比较序列号，再比较站点 ID
        if self.id[1] != other.id[1]:
            return self.id[1] < other.id[1]
        return self.id[0] < other.id[0]


class CRDTEngine:
    """
    CRDT 引擎 - 使用 RGA (Replicated Growable Array) 算法
    实现无冲突的并发文本编辑
    """
    
    def __init__(self, site_id: str = None):
        """初始化 CRDT 引擎
        
        Args:
            site_id: 站点 ID，如果为 None 则自动生成
        """
        self.site_id = site_id or str(uuid.uuid4())[:8]
        self._document: List[CRDTChar] = []
        self._vector_clock: Dict[str, int] = {self.site_id: 0}
        self._pending_ops: List[CRDTOperation] = []
        self._sequence = 0
    
    def local_insert(self, position: int, content: str) -> CRDTOperation:
        """本地插入操作
        
        Args:
            position: 插入位置（可见字符位置）
            content: 插入内容
            
        Returns:
            CRDTOperation 操作对象
        """
        # 递增向量时钟
        self._increment_clock()
        
        # 创建操作
        op = CRDTOperation(
            id=str(uuid.uuid4()),
            type='insert',
            position=position,
            content=content,
            author=self.site_id,
            timestamp=time.time(),
            vector_clock=self._vector_clock.copy()
        )
        
        # 应用到本地文档
        self._apply_insert(position, content)
        
        return op
    
    def local_delete(self, position: int, length: int) -> CRDTOperation:
        """本地删除操作
        
        Args:
            position: 删除起始位置（可见字符位置）
            length: 删除长度
            
        Returns:
            CRDTOperation 操作对象
        """
        # 递增向量时钟
        self._increment_clock()
        
        # 创建操作
        op = CRDTOperation(
            id=str(uuid.uuid4()),
            type='delete',
            position=position,
            length=length,
            author=self.site_id,
            timestamp=time.time(),
            vector_clock=self._vector_clock.copy()
        )
        
        # 应用到本地文档
        self._apply_delete(position, length)
        
        return op
    
    def apply_remote(self, operation: CRDTOperation) -> Tuple[int, str]:
        """应用远程操作
        
        Args:
            operation: 远程操作
            
        Returns:
            (position, content) 实际应用的位置和内容
        """
        # 更新向量时钟
        for site, clock in operation.vector_clock.items():
            self._vector_clock[site] = max(
                self._vector_clock.get(site, 0),
                clock
            )
        
        if operation.type == 'insert':
            # 转换位置（处理并发）
            actual_pos = self._transform_position(operation.position, operation)
            self._apply_insert(actual_pos, operation.content)
            return (actual_pos, operation.content)
        
        elif operation.type == 'delete':
            actual_pos = self._transform_position(operation.position, operation)
            self._apply_delete(actual_pos, operation.length)
            return (actual_pos, '')
        
        return (operation.position, '')
    
    def _apply_insert(self, position: int, content: str) -> None:
        """应用插入操作到文档
        
        Args:
            position: 可见字符位置
            content: 插入内容
        """
        # 找到实际插入位置
        actual_index = self._visible_to_actual(position)
        
        # 插入字符
        for i, char in enumerate(content):
            self._sequence += 1
            crdt_char = CRDTChar(
                char=char,
                id=(self.site_id, self._sequence),
                visible=True
            )
            self._document.insert(actual_index + i, crdt_char)
    
    def _apply_delete(self, position: int, length: int) -> None:
        """应用删除操作到文档（标记为不可见）
        
        Args:
            position: 可见字符位置
            length: 删除长度
        """
        deleted = 0
        visible_pos = 0
        
        for char in self._document:
            if not char.visible:
                continue
            
            if visible_pos >= position and deleted < length:
                char.visible = False
                deleted += 1
            
            visible_pos += 1
            
            if deleted >= length:
                break
    
    def _visible_to_actual(self, visible_pos: int) -> int:
        """将可见位置转换为实际索引
        
        Args:
            visible_pos: 可见字符位置
            
        Returns:
            实际索引
        """
        actual = 0
        visible = 0
        
        for char in self._document:
            if visible >= visible_pos:
                break
            if char.visible:
                visible += 1
            actual += 1
        
        return actual
    
    def _transform_position(self, position: int, operation: CRDTOperation) -> int:
        """转换位置（处理并发操作）
        
        Args:
            position: 原始位置
            operation: 操作
            
        Returns:
            转换后的位置
        """
        # 简化的位置转换
        # 在完整实现中，需要考虑向量时钟来确定操作顺序
        return position
    
    def get_content(self) -> str:
        """获取当前文档内容
        
        Returns:
            可见字符组成的字符串
        """
        return ''.join(char.char for char in self._document if char.visible)
    
    def set_content(self, content: str) -> None:
        """设置文档内容（初始化用）
        
        Args:
            content: 文档内容
        """
        self._document = []
        self._sequence = 0
        
        for char in content:
            self._sequence += 1
            self._document.append(CRDTChar(
                char=char,
                id=(self.site_id, self._sequence),
                visible=True
            ))
    
    def get_state(self) -> bytes:
        """序列化当前状态
        
        Returns:
            序列化的状态数据
        """
        state = {
            'site_id': self.site_id,
            'document': [(c.char, c.id, c.visible) for c in self._document],
            'vector_clock': self._vector_clock,
            'sequence': self._sequence
        }
        return pickle.dumps(state)
    
    def load_state(self, state: bytes) -> None:
        """加载状态
        
        Args:
            state: 序列化的状态数据
        """
        data = pickle.loads(state)
        
        self._document = [
            CRDTChar(char=c[0], id=c[1], visible=c[2])
            for c in data['document']
        ]
        self._vector_clock = data['vector_clock']
        self._sequence = data['sequence']
        
        # 更新本地向量时钟
        if self.site_id not in self._vector_clock:
            self._vector_clock[self.site_id] = 0
    
    def _increment_clock(self) -> Dict[str, int]:
        """递增向量时钟
        
        Returns:
            更新后的向量时钟
        """
        self._vector_clock[self.site_id] = self._vector_clock.get(self.site_id, 0) + 1
        return self._vector_clock.copy()
    
    def get_length(self) -> int:
        """获取可见字符数量
        
        Returns:
            可见字符数量
        """
        return sum(1 for char in self._document if char.visible)
    
    def merge(self, other_state: bytes) -> None:
        """合并另一个状态
        
        Args:
            other_state: 另一个引擎的状态
        """
        other_data = pickle.loads(other_state)
        
        # 合并向量时钟
        for site, clock in other_data['vector_clock'].items():
            self._vector_clock[site] = max(
                self._vector_clock.get(site, 0),
                clock
            )
        
        # 合并文档（简化版本，实际需要更复杂的合并逻辑）
        other_chars = {
            c[1]: CRDTChar(char=c[0], id=c[1], visible=c[2])
            for c in other_data['document']
        }
        
        local_ids = {c.id for c in self._document}
        
        # 添加本地没有的字符
        for char_id, char in other_chars.items():
            if char_id not in local_ids:
                self._document.append(char)
        
        # 重新排序
        self._document.sort()
