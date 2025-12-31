# -*- coding: utf-8 -*-
"""协作网络监控模块"""

import asyncio
import json
import time
import zlib
from dataclasses import dataclass
from typing import Optional, Callable, List
from enum import Enum


class ConnectionQuality(Enum):
    """连接质量等级"""
    EXCELLENT = "excellent"  # < 50ms
    GOOD = "good"           # 50-100ms
    FAIR = "fair"           # 100-200ms
    POOR = "poor"           # 200-500ms
    BAD = "bad"             # > 500ms
    OFFLINE = "offline"


@dataclass
class NetworkStats:
    """网络统计信息"""
    latency: float = 0.0          # 延迟 (ms)
    packet_loss: float = 0.0      # 丢包率 (%)
    bandwidth: float = 0.0        # 带宽 (KB/s)
    bytes_sent: int = 0           # 发送字节数
    bytes_received: int = 0       # 接收字节数
    messages_sent: int = 0        # 发送消息数
    messages_received: int = 0    # 接收消息数
    last_ping_time: float = 0.0   # 最后ping时间
    quality: ConnectionQuality = ConnectionQuality.OFFLINE
    
    def update_quality(self):
        """根据延迟更新连接质量"""
        if self.latency <= 0:
            self.quality = ConnectionQuality.OFFLINE
        elif self.latency < 50:
            self.quality = ConnectionQuality.EXCELLENT
        elif self.latency < 100:
            self.quality = ConnectionQuality.GOOD
        elif self.latency < 200:
            self.quality = ConnectionQuality.FAIR
        elif self.latency < 500:
            self.quality = ConnectionQuality.POOR
        else:
            self.quality = ConnectionQuality.BAD
    
    @property
    def quality_icon(self) -> str:
        """获取质量图标"""
        icons = {
            ConnectionQuality.EXCELLENT: "🟢",
            ConnectionQuality.GOOD: "🟢",
            ConnectionQuality.FAIR: "🟡",
            ConnectionQuality.POOR: "🟠",
            ConnectionQuality.BAD: "🔴",
            ConnectionQuality.OFFLINE: "⚫",
        }
        return icons.get(self.quality, "⚫")
    
    @property
    def quality_text(self) -> str:
        """获取质量文本"""
        texts = {
            ConnectionQuality.EXCELLENT: "极佳",
            ConnectionQuality.GOOD: "良好",
            ConnectionQuality.FAIR: "一般",
            ConnectionQuality.POOR: "较差",
            ConnectionQuality.BAD: "很差",
            ConnectionQuality.OFFLINE: "离线",
        }
        return texts.get(self.quality, "未知")
    
    @property
    def latency_str(self) -> str:
        """格式化延迟"""
        if self.latency <= 0:
            return "--"
        return f"{int(self.latency)}ms"


class NetworkMonitor:
    """网络监控器"""
    
    def __init__(self):
        self.stats = NetworkStats()
        self._ping_history: List[float] = []
        self._max_history = 10
        self._on_quality_change: Optional[Callable] = None
        self._on_disconnect: Optional[Callable] = None
        self._last_quality = ConnectionQuality.OFFLINE
    
    def set_on_quality_change(self, callback: Callable):
        """设置质量变化回调"""
        self._on_quality_change = callback
    
    def set_on_disconnect(self, callback: Callable):
        """设置断线回调"""
        self._on_disconnect = callback
    
    def record_ping(self, latency: float):
        """记录ping延迟"""
        self._ping_history.append(latency)
        if len(self._ping_history) > self._max_history:
            self._ping_history.pop(0)
        
        # 计算平均延迟
        self.stats.latency = sum(self._ping_history) / len(self._ping_history)
        self.stats.last_ping_time = time.time()
        
        # 更新质量
        old_quality = self.stats.quality
        self.stats.update_quality()
        
        # 触发质量变化回调
        if self.stats.quality != old_quality and self._on_quality_change:
            self._on_quality_change(self.stats.quality, old_quality)
    
    def record_sent(self, bytes_count: int):
        """记录发送数据"""
        self.stats.bytes_sent += bytes_count
        self.stats.messages_sent += 1
    
    def record_received(self, bytes_count: int):
        """记录接收数据"""
        self.stats.bytes_received += bytes_count
        self.stats.messages_received += 1
    
    def check_connection(self) -> bool:
        """检查连接状态"""
        if self.stats.last_ping_time == 0:
            return False
        
        # 如果超过30秒没有ping响应，认为断线
        if time.time() - self.stats.last_ping_time > 30:
            self.stats.quality = ConnectionQuality.OFFLINE
            if self._on_disconnect:
                self._on_disconnect()
            return False
        
        return True
    
    def reset(self):
        """重置统计"""
        self.stats = NetworkStats()
        self._ping_history = []


class DataCompressor:
    """数据压缩器"""
    
    @staticmethod
    def compress(data: str) -> bytes:
        """压缩数据"""
        return zlib.compress(data.encode('utf-8'), level=6)
    
    @staticmethod
    def decompress(data: bytes) -> str:
        """解压数据"""
        return zlib.decompress(data).decode('utf-8')
    
    @staticmethod
    def compress_json(obj: dict) -> bytes:
        """压缩JSON对象"""
        json_str = json.dumps(obj, ensure_ascii=False, separators=(',', ':'))
        return DataCompressor.compress(json_str)
    
    @staticmethod
    def decompress_json(data: bytes) -> dict:
        """解压JSON对象"""
        json_str = DataCompressor.decompress(data)
        return json.loads(json_str)


class OfflineCache:
    """离线缓存"""
    
    def __init__(self, max_operations: int = 100):
        self._pending_operations: List[dict] = []
        self._max_operations = max_operations
        self._is_offline = False
    
    @property
    def is_offline(self) -> bool:
        return self._is_offline
    
    @property
    def pending_count(self) -> int:
        return len(self._pending_operations)
    
    def set_offline(self, offline: bool):
        """设置离线状态"""
        self._is_offline = offline
    
    def add_operation(self, operation: dict):
        """添加待同步操作"""
        if len(self._pending_operations) >= self._max_operations:
            # 合并操作或丢弃旧操作
            self._pending_operations.pop(0)
        
        self._pending_operations.append({
            **operation,
            'timestamp': time.time()
        })
    
    def get_pending_operations(self) -> List[dict]:
        """获取待同步操作"""
        return list(self._pending_operations)
    
    def clear_operations(self, count: int = None):
        """清除已同步的操作"""
        if count is None:
            self._pending_operations = []
        else:
            self._pending_operations = self._pending_operations[count:]
    
    def has_pending(self) -> bool:
        """是否有待同步操作"""
        return len(self._pending_operations) > 0


class IncrementalSync:
    """增量同步"""
    
    def __init__(self):
        self._last_content = ""
        self._last_hash = ""
    
    def compute_diff(self, new_content: str) -> Optional[dict]:
        """计算增量差异"""
        if new_content == self._last_content:
            return None
        
        # 简单的行级差异
        old_lines = self._last_content.split('\n')
        new_lines = new_content.split('\n')
        
        changes = []
        
        # 找出变化的行
        max_len = max(len(old_lines), len(new_lines))
        for i in range(max_len):
            old_line = old_lines[i] if i < len(old_lines) else None
            new_line = new_lines[i] if i < len(new_lines) else None
            
            if old_line != new_line:
                changes.append({
                    'line': i,
                    'old': old_line,
                    'new': new_line,
                    'type': 'modify' if old_line and new_line else ('add' if new_line else 'delete')
                })
        
        self._last_content = new_content
        
        return {
            'type': 'incremental',
            'changes': changes,
            'line_count': len(new_lines)
        }
    
    def apply_diff(self, diff: dict) -> str:
        """应用增量差异"""
        lines = self._last_content.split('\n')
        
        for change in diff.get('changes', []):
            line_num = change['line']
            change_type = change['type']
            
            if change_type == 'modify':
                if line_num < len(lines):
                    lines[line_num] = change['new']
            elif change_type == 'add':
                if line_num >= len(lines):
                    lines.append(change['new'])
                else:
                    lines.insert(line_num, change['new'])
            elif change_type == 'delete':
                if line_num < len(lines):
                    lines.pop(line_num)
        
        self._last_content = '\n'.join(lines)
        return self._last_content
    
    def set_content(self, content: str):
        """设置当前内容"""
        self._last_content = content
