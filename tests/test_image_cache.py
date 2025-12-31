# -*- coding: utf-8 -*-
"""
Tests for ImageCache module.

Includes both unit tests and property-based tests.
"""

import unittest
import tempfile
import shutil
import os
from hypothesis import given, strategies as st, settings, assume

import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ui.image_cache import ImageCache


class TestImageCacheUnit(unittest.TestCase):
    """Unit tests for ImageCache."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.cache = ImageCache(cache_dir=self.temp_dir, max_size_mb=1)
    
    def tearDown(self):
        """Clean up test fixtures."""
        try:
            shutil.rmtree(self.temp_dir)
        except Exception:
            pass
    
    def _create_temp_file(self, content: bytes = b'test image data') -> str:
        """Create a temporary file with content."""
        fd, path = tempfile.mkstemp()
        with os.fdopen(fd, 'wb') as f:
            f.write(content)
        return path
    
    def test_put_and_get(self):
        """Test basic put and get operations.
        
        _Requirements: 6.1, 6.2_
        """
        url = 'https://example.com/image.png'
        temp_file = self._create_temp_file()
        
        try:
            cached_path = self.cache.put(url, temp_file)
            
            self.assertIsNotNone(cached_path)
            self.assertTrue(os.path.exists(cached_path))
            
            retrieved = self.cache.get(url)
            self.assertEqual(retrieved, cached_path)
        finally:
            if os.path.exists(temp_file):
                os.remove(temp_file)
    
    def test_get_nonexistent_returns_none(self):
        """Test that get returns None for non-cached URLs."""
        result = self.cache.get('https://example.com/nonexistent.png')
        
        self.assertIsNone(result)
    
    def test_cache_hit_returns_same_path(self):
        """Test that cache hit returns the same path.
        
        _Requirements: 6.2_
        """
        url = 'https://example.com/image.png'
        temp_file = self._create_temp_file()
        
        try:
            cached_path1 = self.cache.put(url, temp_file)
            cached_path2 = self.cache.get(url)
            cached_path3 = self.cache.get(url)
            
            self.assertEqual(cached_path1, cached_path2)
            self.assertEqual(cached_path2, cached_path3)
        finally:
            if os.path.exists(temp_file):
                os.remove(temp_file)
    
    def test_lru_eviction(self):
        """Test LRU eviction when cache exceeds size limit.
        
        _Requirements: 6.3_
        """
        # Create a small cache (1KB)
        small_cache = ImageCache(cache_dir=self.temp_dir, max_size_mb=0.001)
        
        # Add files that exceed the limit
        urls = []
        for i in range(5):
            url = f'https://example.com/image{i}.png'
            urls.append(url)
            temp_file = self._create_temp_file(b'x' * 500)  # 500 bytes each
            try:
                small_cache.put(url, temp_file)
            finally:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
        
        # Oldest entries should be evicted
        # With 1KB limit and 500 byte files, only ~2 should fit
        cached_count = sum(1 for url in urls if small_cache.get(url) is not None)
        self.assertLess(cached_count, 5)
    
    def test_remove(self):
        """Test removing a cached item."""
        url = 'https://example.com/image.png'
        temp_file = self._create_temp_file()
        
        try:
            self.cache.put(url, temp_file)
            self.assertTrue(self.cache.contains(url))
            
            result = self.cache.remove(url)
            
            self.assertTrue(result)
            self.assertFalse(self.cache.contains(url))
            self.assertIsNone(self.cache.get(url))
        finally:
            if os.path.exists(temp_file):
                os.remove(temp_file)
    
    def test_clear(self):
        """Test clearing all cached items."""
        urls = ['https://example.com/image1.png', 'https://example.com/image2.png']
        
        for url in urls:
            temp_file = self._create_temp_file()
            try:
                self.cache.put(url, temp_file)
            finally:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
        
        self.cache.clear()
        
        self.assertEqual(self.cache.count, 0)
        for url in urls:
            self.assertIsNone(self.cache.get(url))
    
    def test_persistence(self):
        """Test cache persistence across instances.
        
        _Requirements: 6.4_
        """
        url = 'https://example.com/image.png'
        temp_file = self._create_temp_file()
        
        try:
            # Put in first cache instance
            self.cache.put(url, temp_file)
            
            # Create new cache instance with same directory
            new_cache = ImageCache(cache_dir=self.temp_dir, max_size_mb=1)
            
            # Should be able to retrieve
            retrieved = new_cache.get(url)
            self.assertIsNotNone(retrieved)
        finally:
            if os.path.exists(temp_file):
                os.remove(temp_file)


class TestImageCacheProperty(unittest.TestCase):
    """
    Property-based tests for ImageCache.
    
    **Feature: code-optimization, Property 8, 9, 10**
    **Validates: Requirements 6.2, 6.3, 6.4**
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
    
    def _create_temp_file(self, content: bytes = b'test') -> str:
        """Create a temporary file."""
        fd, path = tempfile.mkstemp()
        with os.fdopen(fd, 'wb') as f:
            f.write(content)
        return path
    
    @given(st.lists(
        st.text(min_size=5, max_size=30, alphabet=st.characters(whitelist_categories=('L', 'N'))).map(
            lambda x: f'https://example.com/{x}.png'
        ),
        min_size=1,
        max_size=10,
        unique=True
    ))
    @settings(max_examples=100)
    def test_cache_hit_property(self, urls):
        """
        Property 8: 缓存命中
        
        *For any* 已缓存的图片 URL，第二次请求应该直接返回缓存路径，不发起网络请求。
        
        **Feature: code-optimization, Property 8: 缓存命中**
        **Validates: Requirements 6.2**
        """
        cache = ImageCache(cache_dir=self.temp_dir, max_size_mb=100)
        temp_files = []
        
        try:
            # Put all URLs
            cached_paths = {}
            for url in urls:
                temp_file = self._create_temp_file(url.encode())
                temp_files.append(temp_file)
                cached_path = cache.put(url, temp_file)
                if cached_path:
                    cached_paths[url] = cached_path
            
            # Property: get should return the same path for cached URLs
            for url, expected_path in cached_paths.items():
                retrieved = cache.get(url)
                self.assertEqual(retrieved, expected_path,
                    f"Cache hit for '{url}' should return the cached path")
        finally:
            for f in temp_files:
                if os.path.exists(f):
                    os.remove(f)
    
    @given(st.integers(min_value=1, max_value=20))
    @settings(max_examples=50)
    def test_lru_eviction_property(self, num_items):
        """
        Property 9: LRU 缓存淘汰
        
        *For any* 缓存大小超过限制时，最久未使用的条目应该被移除。
        
        **Feature: code-optimization, Property 9: LRU 缓存淘汰**
        **Validates: Requirements 6.3**
        """
        # Create cache with very small size (1KB)
        cache = ImageCache(cache_dir=self.temp_dir, max_size_mb=0.001)
        temp_files = []
        
        try:
            urls = [f'https://example.com/img{i}.png' for i in range(num_items)]
            
            for url in urls:
                temp_file = self._create_temp_file(b'x' * 200)  # 200 bytes each
                temp_files.append(temp_file)
                cache.put(url, temp_file)
            
            # Property: cache size should not exceed limit
            self.assertLessEqual(cache.size, cache.max_size_bytes,
                "Cache size should not exceed the maximum limit")
        finally:
            for f in temp_files:
                if os.path.exists(f):
                    os.remove(f)
    
    @given(st.lists(
        st.text(min_size=5, max_size=20, alphabet=st.characters(whitelist_categories=('L', 'N'))).map(
            lambda x: f'https://example.com/{x}.png'
        ),
        min_size=1,
        max_size=5,
        unique=True
    ))
    @settings(max_examples=50)
    def test_persistence_property(self, urls):
        """
        Property 10: 缓存持久化
        
        *For any* 缓存条目，应用重启后应该能够恢复。
        
        **Feature: code-optimization, Property 10: 缓存持久化**
        **Validates: Requirements 6.4**
        """
        temp_files = []
        
        try:
            # First cache instance
            cache1 = ImageCache(cache_dir=self.temp_dir, max_size_mb=100)
            
            cached_urls = []
            for url in urls:
                temp_file = self._create_temp_file(url.encode())
                temp_files.append(temp_file)
                if cache1.put(url, temp_file):
                    cached_urls.append(url)
            
            # Create new cache instance (simulating restart)
            cache2 = ImageCache(cache_dir=self.temp_dir, max_size_mb=100)
            
            # Property: all cached URLs should be retrievable after "restart"
            for url in cached_urls:
                retrieved = cache2.get(url)
                self.assertIsNotNone(retrieved,
                    f"URL '{url}' should be retrievable after cache restart")
        finally:
            for f in temp_files:
                if os.path.exists(f):
                    os.remove(f)


if __name__ == '__main__':
    unittest.main()
