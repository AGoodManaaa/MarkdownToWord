# -*- coding: utf-8 -*-
"""
Image Downloader - 异步图片下载器

使用线程池实现异步下载，集成缓存。
"""

import os
import tempfile
import threading
import logging
from concurrent.futures import ThreadPoolExecutor, Future
from typing import Callable, Optional, Dict
from urllib.parse import urlparse

logger = logging.getLogger(__name__)


class ImageDownloader:
    """
    异步图片下载器。
    
    使用 ThreadPoolExecutor 实现异步下载，支持缓存集成。
    """
    
    def __init__(self, max_workers: int = 4, cache: Optional['ImageCache'] = None):
        """
        初始化下载器。
        
        Args:
            max_workers: 最大并发下载数
            cache: 图片缓存实例
        """
        self.max_workers = max_workers
        self.executor = ThreadPoolExecutor(max_workers=max_workers)
        self.cache = cache
        self._pending: Dict[str, Future] = {}
        self._lock = threading.Lock()
        self._shutdown = False
    
    def download_async(
        self,
        url: str,
        on_complete: Callable[[str, Optional[str]], None],
        on_error: Optional[Callable[[str, Exception], None]] = None
    ) -> None:
        """
        异步下载图片。
        
        此方法立即返回，不阻塞调用线程。
        
        Args:
            url: 图片 URL
            on_complete: 下载完成回调 (url, local_path)
            on_error: 下载失败回调 (url, exception)
        """
        if self._shutdown:
            if on_error:
                on_error(url, RuntimeError("Downloader is shutdown"))
            return
        
        # 检查缓存
        if self.cache:
            cached_path = self.cache.get(url)
            if cached_path:
                on_complete(url, cached_path)
                return
        
        # 检查是否已在下载中
        with self._lock:
            if url in self._pending:
                return  # 避免重复下载
            
            future = self.executor.submit(self._download, url)
            self._pending[url] = future
        
        def callback(f: Future):
            with self._lock:
                self._pending.pop(url, None)
            
            try:
                local_path = f.result()
                if local_path and self.cache:
                    cached_path = self.cache.put(url, local_path)
                    # 删除临时文件
                    if cached_path and cached_path != local_path:
                        try:
                            os.remove(local_path)
                        except OSError:
                            pass
                    local_path = cached_path or local_path
                on_complete(url, local_path)
            except Exception as e:
                logger.error(f"Download failed for {url}: {e}")
                if on_error:
                    on_error(url, e)
                else:
                    on_complete(url, None)
        
        future.add_done_callback(callback)
    
    def _download(self, url: str) -> Optional[str]:
        """
        实际下载逻辑（在线程中执行）。
        
        Args:
            url: 图片 URL
            
        Returns:
            下载的临时文件路径
        """
        import requests
        
        try:
            response = requests.get(url, timeout=30, stream=True)
            response.raise_for_status()
            
            ext = self._get_extension(url, response.headers.get('Content-Type', ''))
            fd, temp_path = tempfile.mkstemp(suffix=ext)
            
            try:
                with os.fdopen(fd, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)
            except Exception:
                os.close(fd)
                raise
            
            return temp_path
            
        except Exception as e:
            logger.error(f"Failed to download {url}: {e}")
            raise
    
    def _get_extension(self, url: str, content_type: str) -> str:
        """
        获取文件扩展名。
        
        Args:
            url: 图片 URL
            content_type: Content-Type 头
            
        Returns:
            文件扩展名
        """
        # 从 URL 获取扩展名
        parsed = urlparse(url)
        path = parsed.path
        ext = os.path.splitext(path)[1].lower()
        
        if ext in ('.jpg', '.jpeg', '.png', '.gif', '.webp', '.bmp', '.svg'):
            return ext
        
        # 从 Content-Type 获取
        content_type_map = {
            'image/jpeg': '.jpg',
            'image/png': '.png',
            'image/gif': '.gif',
            'image/webp': '.webp',
            'image/bmp': '.bmp',
            'image/svg+xml': '.svg',
        }
        
        for ct, extension in content_type_map.items():
            if ct in content_type:
                return extension
        
        return '.img'
    
    def is_downloading(self, url: str) -> bool:
        """检查 URL 是否正在下载中。"""
        with self._lock:
            return url in self._pending
    
    def get_pending_count(self) -> int:
        """获取正在下载的数量。"""
        with self._lock:
            return len(self._pending)
    
    def get_active_thread_count(self) -> int:
        """获取活跃线程数。"""
        # ThreadPoolExecutor 没有直接暴露活跃线程数
        # 使用 pending 数量作为近似值，但不超过 max_workers
        with self._lock:
            return min(len(self._pending), self.max_workers)
    
    def shutdown(self, wait: bool = False) -> None:
        """
        关闭下载器。
        
        Args:
            wait: 是否等待所有下载完成
        """
        self._shutdown = True
        self.executor.shutdown(wait=wait)
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.shutdown(wait=True)
        return False
