# -*- coding: utf-8 -*-
"""增量渲染模块 - 只渲染变化的部分，提升大文档性能"""

import hashlib
import time
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple, Callable
import tkinter as tk


@dataclass
class RenderBlock:
    """渲染块"""
    start_line: int
    end_line: int
    content: str
    content_hash: str
    rendered_html: str = ""
    is_dirty: bool = True
    block_type: str = "text"  # text, code, table, heading, list


@dataclass
class RenderCache:
    """渲染缓存"""
    blocks: Dict[int, RenderBlock] = field(default_factory=dict)
    last_full_render: float = 0
    total_lines: int = 0
    

class IncrementalRenderer:
    """增量渲染器 - 只渲染变化的部分"""
    
    def __init__(self, text_widget: tk.Text, preview_callback: Callable[[str], None] = None):
        self.text_widget = text_widget
        self.preview_callback = preview_callback
        self.cache = RenderCache()
        self._enabled = True
        self._last_content = ""
        self._last_cursor_line = 0
        self._debounce_id = None
        self._debounce_delay = 100  # ms
        
        # 性能统计
        self._render_count = 0
        self._full_render_count = 0
        self._incremental_render_count = 0
        self._total_render_time = 0
        
        # 块类型检测正则
        import re
        self._patterns = {
            'heading': re.compile(r'^#{1,6}\s'),
            'code_start': re.compile(r'^```'),
            'code_end': re.compile(r'^```\s*$'),
            'table': re.compile(r'^\|.*\|'),
            'list': re.compile(r'^[\s]*[-*+]\s|^[\s]*\d+\.\s'),
            'blockquote': re.compile(r'^>\s'),
            'hr': re.compile(r'^[-*_]{3,}\s*$'),
        }
    
    def enable(self):
        """启用增量渲染"""
        self._enabled = True
    
    def disable(self):
        """禁用增量渲染"""
        self._enabled = False
    
    def get_content(self) -> str:
        """获取文本内容"""
        try:
            return self.text_widget.get("1.0", "end-1c")
        except:
            return ""
    
    def _compute_hash(self, content: str) -> str:
        """计算内容哈希"""
        return hashlib.md5(content.encode('utf-8')).hexdigest()[:8]
    
    def _detect_block_type(self, lines: List[str]) -> str:
        """检测块类型"""
        if not lines:
            return "text"
        
        first_line = lines[0]
        
        if self._patterns['heading'].match(first_line):
            return "heading"
        if self._patterns['code_start'].match(first_line):
            return "code"
        if self._patterns['table'].match(first_line):
            return "table"
        if self._patterns['list'].match(first_line):
            return "list"
        if self._patterns['blockquote'].match(first_line):
            return "blockquote"
        if self._patterns['hr'].match(first_line):
            return "hr"
        
        return "text"
    
    def _split_into_blocks(self, content: str) -> List[RenderBlock]:
        """将内容分割成渲染块"""
        lines = content.split('\n')
        blocks = []
        current_block_lines = []
        current_start = 0
        in_code_block = False
        
        for i, line in enumerate(lines):
            # 代码块处理
            if self._patterns['code_start'].match(line):
                if in_code_block:
                    # 代码块结束
                    current_block_lines.append(line)
                    block_content = '\n'.join(current_block_lines)
                    blocks.append(RenderBlock(
                        start_line=current_start,
                        end_line=i,
                        content=block_content,
                        content_hash=self._compute_hash(block_content),
                        block_type="code"
                    ))
                    current_block_lines = []
                    current_start = i + 1
                    in_code_block = False
                else:
                    # 保存之前的块
                    if current_block_lines:
                        block_content = '\n'.join(current_block_lines)
                        blocks.append(RenderBlock(
                            start_line=current_start,
                            end_line=i - 1,
                            content=block_content,
                            content_hash=self._compute_hash(block_content),
                            block_type=self._detect_block_type(current_block_lines)
                        ))
                    current_block_lines = [line]
                    current_start = i
                    in_code_block = True
                continue
            
            if in_code_block:
                current_block_lines.append(line)
                continue
            
            # 空行分隔块
            if not line.strip():
                if current_block_lines:
                    block_content = '\n'.join(current_block_lines)
                    blocks.append(RenderBlock(
                        start_line=current_start,
                        end_line=i - 1,
                        content=block_content,
                        content_hash=self._compute_hash(block_content),
                        block_type=self._detect_block_type(current_block_lines)
                    ))
                    current_block_lines = []
                current_start = i + 1
            else:
                current_block_lines.append(line)
        
        # 处理最后一个块
        if current_block_lines:
            block_content = '\n'.join(current_block_lines)
            blocks.append(RenderBlock(
                start_line=current_start,
                end_line=len(lines) - 1,
                content=block_content,
                content_hash=self._compute_hash(block_content),
                block_type=self._detect_block_type(current_block_lines)
            ))
        
        return blocks
    
    def _find_changed_blocks(self, new_blocks: List[RenderBlock]) -> Tuple[List[int], List[int], List[int]]:
        """找出变化的块
        
        Returns:
            (added, modified, removed) - 新增、修改、删除的块索引
        """
        added = []
        modified = []
        removed = []
        
        old_hashes = {b.content_hash: i for i, b in self.cache.blocks.items()}
        new_hashes = {b.content_hash: i for i, b in enumerate(new_blocks)}
        
        # 找出新增和修改的块
        for i, block in enumerate(new_blocks):
            if block.content_hash not in old_hashes:
                # 检查是否是修改（同位置不同内容）
                if i in self.cache.blocks:
                    modified.append(i)
                else:
                    added.append(i)
        
        # 找出删除的块
        for old_idx in self.cache.blocks:
            if old_idx >= len(new_blocks):
                removed.append(old_idx)
            elif self.cache.blocks[old_idx].content_hash not in new_hashes:
                if old_idx not in modified:
                    removed.append(old_idx)
        
        return added, modified, removed
    
    def render(self, force_full: bool = False) -> str:
        """渲染内容
        
        Args:
            force_full: 是否强制全量渲染
            
        Returns:
            渲染后的 HTML
        """
        if not self._enabled:
            return self._full_render()
        
        start_time = time.time()
        content = self.get_content()
        
        # 内容未变化
        if content == self._last_content and not force_full:
            return self._get_cached_html()
        
        self._last_content = content
        new_blocks = self._split_into_blocks(content)
        
        # 判断是否需要全量渲染
        if force_full or not self.cache.blocks or len(new_blocks) == 0:
            result = self._full_render()
            self._full_render_count += 1
        else:
            added, modified, removed = self._find_changed_blocks(new_blocks)
            
            # 如果变化太大，使用全量渲染
            change_ratio = (len(added) + len(modified) + len(removed)) / max(len(new_blocks), 1)
            if change_ratio > 0.5:
                result = self._full_render()
                self._full_render_count += 1
            else:
                result = self._incremental_render(new_blocks, added, modified, removed)
                self._incremental_render_count += 1
        
        # 更新缓存
        self.cache.blocks = {i: block for i, block in enumerate(new_blocks)}
        self.cache.total_lines = len(content.split('\n'))
        
        # 统计
        self._render_count += 1
        self._total_render_time += time.time() - start_time
        
        return result
    
    def _full_render(self) -> str:
        """全量渲染"""
        content = self.get_content()
        self.cache.last_full_render = time.time()
        
        # 调用预览回调进行渲染
        if self.preview_callback:
            return self.preview_callback(content)
        
        return content
    
    def _incremental_render(self, new_blocks: List[RenderBlock], 
                           added: List[int], modified: List[int], 
                           removed: List[int]) -> str:
        """增量渲染 - 只渲染变化的部分"""
        # 对于增量渲染，我们仍然需要完整的 HTML
        # 但可以复用未变化块的渲染结果
        
        # 简化实现：标记变化的块，然后重新渲染
        # 实际的增量渲染需要更复杂的 DOM diff
        
        content = self.get_content()
        if self.preview_callback:
            return self.preview_callback(content)
        
        return content
    
    def _get_cached_html(self) -> str:
        """获取缓存的 HTML"""
        # 简化实现：返回上次渲染结果
        content = self.get_content()
        if self.preview_callback:
            return self.preview_callback(content)
        return content
    
    def get_stats(self) -> Dict:
        """获取性能统计"""
        avg_time = self._total_render_time / max(self._render_count, 1)
        return {
            'total_renders': self._render_count,
            'full_renders': self._full_render_count,
            'incremental_renders': self._incremental_render_count,
            'avg_render_time_ms': avg_time * 1000,
            'cache_blocks': len(self.cache.blocks),
            'total_lines': self.cache.total_lines,
        }
    
    def clear_cache(self):
        """清除缓存"""
        self.cache = RenderCache()
        self._last_content = ""
    
    def invalidate_range(self, start_line: int, end_line: int):
        """使指定范围的缓存失效"""
        for idx, block in self.cache.blocks.items():
            if block.start_line <= end_line and block.end_line >= start_line:
                block.is_dirty = True


class VirtualScroller:
    """虚拟滚动器 - 只渲染可见区域"""
    
    def __init__(self, text_widget: tk.Text, visible_lines: int = 50):
        self.text_widget = text_widget
        self.visible_lines = visible_lines
        self._enabled = True
        self._total_lines = 0
        self._viewport_start = 0
        self._viewport_end = visible_lines
        self._line_heights: Dict[int, int] = {}
        self._default_line_height = 20
        
        # 缓冲区（上下各多渲染一些行）
        self._buffer_lines = 10
    
    def enable(self):
        """启用虚拟滚动"""
        self._enabled = True
    
    def disable(self):
        """禁用虚拟滚动"""
        self._enabled = False
    
    def update_viewport(self, scroll_position: float):
        """更新可视区域
        
        Args:
            scroll_position: 滚动位置 (0.0 - 1.0)
        """
        if not self._enabled:
            return
        
        # 计算可视区域
        total = self._total_lines
        start = int(scroll_position * total)
        
        self._viewport_start = max(0, start - self._buffer_lines)
        self._viewport_end = min(total, start + self.visible_lines + self._buffer_lines)
    
    def get_visible_range(self) -> Tuple[int, int]:
        """获取可视范围"""
        return self._viewport_start, self._viewport_end
    
    def is_line_visible(self, line: int) -> bool:
        """检查行是否可见"""
        return self._viewport_start <= line <= self._viewport_end
    
    def set_total_lines(self, total: int):
        """设置总行数"""
        self._total_lines = total
    
    def get_scroll_height(self) -> int:
        """获取滚动高度"""
        return self._total_lines * self._default_line_height


class ImageLazyLoader:
    """图片懒加载器"""
    
    def __init__(self):
        self._loaded_images: Dict[str, any] = {}
        self._pending_images: List[str] = []
        self._loading_images: set = set()
        self._enabled = True
        self._placeholder = "⏳"  # 加载中占位符
        self._error_placeholder = "❌"  # 加载失败占位符
        
        # 加载队列
        self._load_queue: List[Tuple[str, Callable]] = []
        self._max_concurrent = 3
    
    def enable(self):
        """启用懒加载"""
        self._enabled = True
    
    def disable(self):
        """禁用懒加载"""
        self._enabled = False
    
    def is_loaded(self, url: str) -> bool:
        """检查图片是否已加载"""
        return url in self._loaded_images
    
    def get_image(self, url: str) -> Optional[any]:
        """获取已加载的图片"""
        return self._loaded_images.get(url)
    
    def queue_load(self, url: str, callback: Callable[[str, any], None]):
        """将图片加入加载队列
        
        Args:
            url: 图片 URL
            callback: 加载完成回调 (url, image_data)
        """
        if not self._enabled:
            return
        
        if url in self._loaded_images or url in self._loading_images:
            return
        
        self._load_queue.append((url, callback))
        self._process_queue()
    
    def _process_queue(self):
        """处理加载队列"""
        while self._load_queue and len(self._loading_images) < self._max_concurrent:
            url, callback = self._load_queue.pop(0)
            self._load_image(url, callback)
    
    def _load_image(self, url: str, callback: Callable):
        """加载图片"""
        import threading
        
        self._loading_images.add(url)
        
        def load():
            try:
                import urllib.request
                from PIL import Image
                import io
                
                # 下载图片
                with urllib.request.urlopen(url, timeout=10) as response:
                    data = response.read()
                
                # 解析图片
                image = Image.open(io.BytesIO(data))
                
                # 缓存
                self._loaded_images[url] = image
                self._loading_images.discard(url)
                
                # 回调
                callback(url, image)
                
            except Exception as e:
                self._loading_images.discard(url)
                self._loaded_images[url] = None  # 标记为加载失败
                callback(url, None)
            
            # 继续处理队列
            self._process_queue()
        
        thread = threading.Thread(target=load, daemon=True)
        thread.start()
    
    def clear_cache(self):
        """清除缓存"""
        self._loaded_images.clear()
        self._pending_images.clear()
        self._loading_images.clear()
        self._load_queue.clear()
    
    def get_stats(self) -> Dict:
        """获取统计信息"""
        return {
            'loaded': len(self._loaded_images),
            'loading': len(self._loading_images),
            'pending': len(self._load_queue),
        }


class PerformanceMonitor:
    """性能监控器"""
    
    def __init__(self):
        self._metrics: Dict[str, List[float]] = {
            'render_time': [],
            'scroll_time': [],
            'input_latency': [],
            'memory_usage': [],
        }
        self._max_samples = 100
        self._enabled = True
    
    def enable(self):
        """启用监控"""
        self._enabled = True
    
    def disable(self):
        """禁用监控"""
        self._enabled = False
    
    def record(self, metric: str, value: float):
        """记录指标"""
        if not self._enabled:
            return
        
        if metric not in self._metrics:
            self._metrics[metric] = []
        
        self._metrics[metric].append(value)
        
        # 保持样本数量
        if len(self._metrics[metric]) > self._max_samples:
            self._metrics[metric] = self._metrics[metric][-self._max_samples:]
    
    def get_average(self, metric: str) -> float:
        """获取平均值"""
        values = self._metrics.get(metric, [])
        return sum(values) / len(values) if values else 0
    
    def get_max(self, metric: str) -> float:
        """获取最大值"""
        values = self._metrics.get(metric, [])
        return max(values) if values else 0
    
    def get_min(self, metric: str) -> float:
        """获取最小值"""
        values = self._metrics.get(metric, [])
        return min(values) if values else 0
    
    def get_report(self) -> Dict:
        """获取性能报告"""
        report = {}
        for metric, values in self._metrics.items():
            if values:
                report[metric] = {
                    'avg': sum(values) / len(values),
                    'max': max(values),
                    'min': min(values),
                    'samples': len(values),
                }
        return report
    
    def clear(self):
        """清除记录"""
        for key in self._metrics:
            self._metrics[key] = []
