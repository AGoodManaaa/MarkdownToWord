# -*- coding: utf-8 -*-
"""
Image Cache - 图片缓存管理

实现 LRU 缓存策略，支持持久化。
"""

import os
import json
import hashlib
import shutil
import logging
from collections import OrderedDict
from typing import Optional, Dict, Any

logger = logging.getLogger(__name__)


class ImageCache:
    """
    图片缓存管理（LRU 策略）。
    
    缓存网络图片到本地，支持大小限制和持久化。
    """
    
    def __init__(self, cache_dir: Optional[str] = None, max_size_mb: int = 100):
        """
        初始化图片缓存。
        
        Args:
            cache_dir: 缓存目录路径，None 表示使用默认目录
            max_size_mb: 最大缓存大小（MB）
        """
        self.cache_dir = cache_dir or os.path.join(
            os.path.expanduser('~'), '.md2word_cache', 'images'
        )
        self.max_size_bytes = max_size_mb * 1024 * 1024
        self.index_file = os.path.join(self.cache_dir, 'index.json')
        self._index: OrderedDict[str, Dict[str, Any]] = OrderedDict()
        self._current_size = 0
        
        self._ensure_cache_dir()
        self._load_index()
    
    def _ensure_cache_dir(self) -> None:
        """确保缓存目录存在。"""
        try:
            os.makedirs(self.cache_dir, exist_ok=True)
        except OSError as e:
            logger.error(f"Failed to create cache directory: {e}")
    
    def get(self, url: str) -> Optional[str]:
        """
        获取缓存的图片路径。
        
        Args:
            url: 图片 URL
            
        Returns:
            本地缓存路径，如果不存在则返回 None
        """
        key = self._url_to_key(url)
        if key not in self._index:
            return None
        
        entry = self._index[key]
        local_path = entry['path']
        
        # 验证文件存在
        if not os.path.exists(local_path):
            del self._index[key]
            self._current_size -= entry.get('size', 0)
            self._save_index()
            return None
        
        # LRU: 移到末尾（最近使用）
        self._index.move_to_end(key)
        return local_path
    
    def put(self, url: str, local_path: str) -> Optional[str]:
        """
        添加图片到缓存。
        
        Args:
            url: 图片 URL
            local_path: 临时文件路径
            
        Returns:
            缓存文件路径，如果失败则返回 None
        """
        if not os.path.exists(local_path):
            return None
        
        key = self._url_to_key(url)
        
        # 如果已存在，先删除旧条目
        if key in self._index:
            self._remove_entry(key)
        
        # 复制到缓存目录
        ext = os.path.splitext(local_path)[1] or '.img'
        cache_path = os.path.join(self.cache_dir, f"{key}{ext}")
        
        try:
            shutil.copy2(local_path, cache_path)
        except (IOError, OSError) as e:
            logger.error(f"Failed to copy file to cache: {e}")
            return None
        
        file_size = os.path.getsize(cache_path)
        
        self._index[key] = {
            'url': url,
            'path': cache_path,
            'size': file_size
        }
        self._current_size += file_size
        
        # 检查是否需要清理
        self._evict_if_needed()
        self._save_index()
        
        return cache_path
    
    def remove(self, url: str) -> bool:
        """
        从缓存中移除图片。
        
        Args:
            url: 图片 URL
            
        Returns:
            True 如果成功移除，否则 False
        """
        key = self._url_to_key(url)
        if key not in self._index:
            return False
        
        self._remove_entry(key)
        self._save_index()
        return True
    
    def clear(self) -> None:
        """清空所有缓存。"""
        for key in list(self._index.keys()):
            self._remove_entry(key)
        self._save_index()
    
    def _remove_entry(self, key: str) -> None:
        """移除单个缓存条目。"""
        if key not in self._index:
            return
        
        entry = self._index[key]
        try:
            if os.path.exists(entry['path']):
                os.remove(entry['path'])
        except OSError:
            pass
        
        self._current_size -= entry.get('size', 0)
        del self._index[key]
    
    def _evict_if_needed(self) -> None:
        """LRU 清理：移除最久未使用的条目直到大小符合限制。"""
        while self._current_size > self.max_size_bytes and self._index:
            # popitem(last=False) 移除最早的条目（LRU）
            key, entry = self._index.popitem(last=False)
            try:
                if os.path.exists(entry['path']):
                    os.remove(entry['path'])
            except OSError:
                pass
            self._current_size -= entry.get('size', 0)
    
    def _url_to_key(self, url: str) -> str:
        """URL 转缓存 key。"""
        return hashlib.md5(url.encode()).hexdigest()
    
    def _load_index(self) -> None:
        """加载缓存索引。"""
        if not os.path.exists(self.index_file):
            return
        
        try:
            with open(self.index_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                entries = data.get('entries', [])
                self._index = OrderedDict(entries)
                self._current_size = sum(
                    e.get('size', 0) for e in self._index.values()
                )
        except (json.JSONDecodeError, IOError) as e:
            logger.warning(f"Failed to load cache index: {e}")
            self._index = OrderedDict()
            self._current_size = 0
    
    def _save_index(self) -> None:
        """保存缓存索引。"""
        try:
            with open(self.index_file, 'w', encoding='utf-8') as f:
                json.dump({'entries': list(self._index.items())}, f)
        except IOError as e:
            logger.error(f"Failed to save cache index: {e}")
    
    @property
    def size(self) -> int:
        """当前缓存大小（字节）。"""
        return self._current_size
    
    @property
    def count(self) -> int:
        """缓存条目数量。"""
        return len(self._index)
    
    def contains(self, url: str) -> bool:
        """检查 URL 是否在缓存中。"""
        key = self._url_to_key(url)
        return key in self._index
