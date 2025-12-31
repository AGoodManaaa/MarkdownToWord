# -*- coding: utf-8 -*-
"""
Tests for ImageDownloader module.

Includes both unit tests and property-based tests.
"""

import unittest
import tempfile
import shutil
import os
import time
import threading
from unittest.mock import MagicMock, patch
from hypothesis import given, strategies as st, settings

import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ui.image_downloader import ImageDownloader
from ui.image_cache import ImageCache


class TestImageDownloaderUnit(unittest.TestCase):
    """Unit tests for ImageDownloader."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.cache = ImageCache(cache_dir=self.temp_dir, max_size_mb=10)
        self.downloader = ImageDownloader(max_workers=2, cache=self.cache)
    
    def tearDown(self):
        """Clean up test fixtures."""
        self.downloader.shutdown(wait=True)
        try:
            shutil.rmtree(self.temp_dir)
        except Exception:
            pass
    
    def test_download_async_returns_immediately(self):
        """Test that download_async returns immediately (non-blocking).
        
        _Requirements: 5.1, 5.3_
        """
        url = 'https://example.com/image.png'
        callback = MagicMock()
        
        start_time = time.time()
        self.downloader.download_async(url, callback)
        elapsed = time.time() - start_time
        
        # Should return almost immediately (< 100ms)
        self.assertLess(elapsed, 0.1,
            "download_async should return immediately without blocking")
    
    def test_download_async_uses_cache(self):
        """Test that download_async uses cache when available.
        
        _Requirements: 6.1, 6.2_
        """
        url = 'https://example.com/cached.png'
        
        # Pre-populate cache
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
        temp_file.write(b'cached image')
        temp_file.close()
        
        try:
            cached_path = self.cache.put(url, temp_file.name)
            
            callback = MagicMock()
            self.downloader.download_async(url, callback)
            
            # Should call callback immediately with cached path
            time.sleep(0.1)
            callback.assert_called_once_with(url, cached_path)
        finally:
            if os.path.exists(temp_file.name):
                os.remove(temp_file.name)
    
    def test_is_downloading(self):
        """Test is_downloading method."""
        url = 'https://example.com/image.png'
        
        # Mock the download to be slow
        with patch.object(self.downloader, '_download') as mock_download:
            mock_download.side_effect = lambda u: time.sleep(1) or '/tmp/test.png'
            
            self.downloader.download_async(url, MagicMock())
            time.sleep(0.05)  # Give time for thread to start
            
            self.assertTrue(self.downloader.is_downloading(url))
    
    def test_get_pending_count(self):
        """Test get_pending_count method."""
        self.assertEqual(self.downloader.get_pending_count(), 0)
    
    def test_shutdown(self):
        """Test shutdown method."""
        self.downloader.shutdown(wait=True)
        
        # After shutdown, new downloads should fail
        callback = MagicMock()
        error_callback = MagicMock()
        self.downloader.download_async('https://example.com/test.png', callback, error_callback)
        
        time.sleep(0.1)
        error_callback.assert_called_once()
    
    def test_get_extension_from_url(self):
        """Test _get_extension extracts extension from URL."""
        ext = self.downloader._get_extension('https://example.com/image.png', '')
        self.assertEqual(ext, '.png')
        
        ext = self.downloader._get_extension('https://example.com/photo.jpg', '')
        self.assertEqual(ext, '.jpg')
    
    def test_get_extension_from_content_type(self):
        """Test _get_extension uses Content-Type when URL has no extension."""
        ext = self.downloader._get_extension('https://example.com/image', 'image/png')
        self.assertEqual(ext, '.png')
        
        ext = self.downloader._get_extension('https://example.com/image', 'image/jpeg')
        self.assertEqual(ext, '.jpg')


class TestImageDownloaderProperty(unittest.TestCase):
    """
    Property-based tests for ImageDownloader.
    
    **Feature: code-optimization, Property 6, 7**
    **Validates: Requirements 5.1, 5.3, 5.4**
    """
    
    def setUp(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
    
    def tearDown(self):
        """Clean up test fixtures."""
        try:
            shutil.rmtree(self.temp_dir)
        except Exception:
            pass
    
    @given(st.integers(min_value=1, max_value=10))
    @settings(max_examples=50, deadline=None)
    def test_async_download_non_blocking_property(self, num_downloads):
        """
        Property 6: 异步下载非阻塞
        
        *For any* 图片下载请求，download_async 应该立即返回，不阻塞调用线程。
        
        **Feature: code-optimization, Property 6: 异步下载非阻塞**
        **Validates: Requirements 5.1, 5.3**
        """
        downloader = ImageDownloader(max_workers=2)
        
        try:
            urls = [f'https://example.com/image{i}.png' for i in range(num_downloads)]
            
            start_time = time.time()
            for url in urls:
                downloader.download_async(url, MagicMock())
            elapsed = time.time() - start_time
            
            # Property: all download_async calls should complete quickly
            # Allow up to 500ms for thread pool initialization overhead
            self.assertLess(elapsed, 0.5,
                f"Initiating {num_downloads} downloads should not block (took {elapsed:.3f}s)")
        finally:
            downloader.shutdown(wait=False)
    
    @given(st.integers(min_value=1, max_value=8))
    @settings(max_examples=50, deadline=None)
    def test_concurrent_download_limit_property(self, max_workers):
        """
        Property 7: 并发下载限制
        
        *For any* 并发下载请求数量，活跃线程数不应超过 max_workers。
        
        **Feature: code-optimization, Property 7: 并发下载限制**
        **Validates: Requirements 5.4**
        """
        downloader = ImageDownloader(max_workers=max_workers)
        active_count = []
        lock = threading.Lock()
        
        def slow_download(url):
            with lock:
                active_count.append(downloader.get_active_thread_count())
            time.sleep(0.1)
            return None
        
        try:
            with patch.object(downloader, '_download', side_effect=slow_download):
                # Start more downloads than max_workers
                num_downloads = max_workers * 2
                for i in range(num_downloads):
                    downloader.download_async(
                        f'https://example.com/img{i}.png',
                        MagicMock()
                    )
                
                time.sleep(0.2)  # Wait for some downloads to start
                
                # Property: active thread count should never exceed max_workers
                for count in active_count:
                    self.assertLessEqual(count, max_workers,
                        f"Active thread count {count} should not exceed max_workers {max_workers}")
        finally:
            downloader.shutdown(wait=False)


if __name__ == '__main__':
    unittest.main()
