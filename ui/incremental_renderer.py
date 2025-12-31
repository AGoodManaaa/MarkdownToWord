# -*- coding: utf-8 -*-
"""
Incremental Preview Renderer - 增量预览渲染器

只渲染变化的部分，提高大文档的预览性能。
"""

import hashlib
import logging
from typing import List, Dict, Any, Optional, Tuple

logger = logging.getLogger(__name__)


class IncrementalPreviewRenderer:
    """
    增量预览渲染器。
    
    通过块级差异计算，只更新变化的部分，提高渲染性能。
    """
    
    def __init__(self, preview_widget, change_threshold: float = 0.3):
        """
        初始化增量渲染器。
        
        Args:
            preview_widget: 预览组件
            change_threshold: 变化阈值，超过此比例使用全量渲染
        """
        self.preview = preview_widget
        self.change_threshold = change_threshold
        self._last_content = ""
        self._last_blocks: List[str] = []
        self._last_hashes: List[str] = []
    
    def render(self, content: str) -> Dict[str, Any]:
        """
        渲染内容（自动选择增量或全量）。
        
        Args:
            content: Markdown 内容
            
        Returns:
            渲染结果信息，包含 type ('full' 或 'incremental') 和 changes
        """
        if not content:
            result = self._full_render("")
            return {'type': 'full', 'changes': [], 'rendered': True}
        
        # 解析为块
        new_blocks = self._parse_blocks(content)
        new_hashes = [self._hash_block(b) for b in new_blocks]
        
        # 计算差异
        diff = self._compute_diff(self._last_hashes, new_hashes, self._last_blocks, new_blocks)
        
        if diff['type'] == 'full':
            self._full_render(content)
        else:
            self._incremental_render(diff, new_blocks)
        
        # 更新状态
        self._last_content = content
        self._last_blocks = new_blocks
        self._last_hashes = new_hashes
        
        return diff
    
    def _parse_blocks(self, content: str) -> List[str]:
        """
        将内容解析为块。
        
        使用空行作为块分隔符。
        
        Args:
            content: Markdown 内容
            
        Returns:
            块列表
        """
        if not content:
            return []
        
        blocks = []
        current_block = []
        in_code_block = False
        
        for line in content.split('\n'):
            # 检测代码块
            if line.strip().startswith('```'):
                in_code_block = not in_code_block
            
            if not in_code_block and line.strip() == '' and current_block:
                blocks.append('\n'.join(current_block))
                current_block = []
            else:
                current_block.append(line)
        
        if current_block:
            blocks.append('\n'.join(current_block))
        
        return blocks
    
    def _hash_block(self, block: str) -> str:
        """计算块的哈希值。"""
        return hashlib.md5(block.encode()).hexdigest()
    
    def _compute_diff(
        self, 
        old_hashes: List[str], 
        new_hashes: List[str],
        old_blocks: List[str],
        new_blocks: List[str]
    ) -> Dict[str, Any]:
        """
        计算块级差异。
        
        Args:
            old_hashes: 旧块哈希列表
            new_hashes: 新块哈希列表
            old_blocks: 旧块列表
            new_blocks: 新块列表
            
        Returns:
            差异信息
        """
        # 如果没有旧内容，使用全量渲染
        if not old_hashes:
            return {'type': 'full'}
        
        # 计算变化比例
        changed_ratio = self._calculate_change_ratio(old_hashes, new_hashes)
        
        # 如果变化超过阈值，使用全量渲染
        if changed_ratio > self.change_threshold:
            return {'type': 'full'}
        
        # 找出变化的块
        changes = []
        max_len = max(len(old_hashes), len(new_hashes))
        
        for i in range(max_len):
            old_hash = old_hashes[i] if i < len(old_hashes) else None
            new_hash = new_hashes[i] if i < len(new_hashes) else None
            
            if old_hash != new_hash:
                change = {
                    'index': i,
                    'old_hash': old_hash,
                    'new_hash': new_hash,
                    'old': old_blocks[i] if i < len(old_blocks) else None,
                    'new': new_blocks[i] if i < len(new_blocks) else None,
                }
                
                if old_hash is None:
                    change['action'] = 'insert'
                elif new_hash is None:
                    change['action'] = 'delete'
                else:
                    change['action'] = 'update'
                
                changes.append(change)
        
        return {'type': 'incremental', 'changes': changes}
    
    def _calculate_change_ratio(self, old_hashes: List[str], new_hashes: List[str]) -> float:
        """
        计算变化比例。
        
        Args:
            old_hashes: 旧块哈希列表
            new_hashes: 新块哈希列表
            
        Returns:
            变化比例 (0.0 - 1.0)
        """
        if not old_hashes and not new_hashes:
            return 0.0
        
        # 计算不同的块数量（按位置比较）
        max_len = max(len(old_hashes), len(new_hashes))
        min_len = min(len(old_hashes), len(new_hashes))
        
        different = 0
        for i in range(min_len):
            if old_hashes[i] != new_hashes[i]:
                different += 1
        
        # 长度差异也算作变化
        different += abs(len(old_hashes) - len(new_hashes))
        
        return different / max_len if max_len > 0 else 0.0
    
    def _full_render(self, content: str) -> None:
        """
        全量渲染。
        
        Args:
            content: Markdown 内容
        """
        if hasattr(self.preview, 'update_preview'):
            self.preview.update_preview(content)
        elif hasattr(self.preview, 'set_content'):
            self.preview.set_content(content)
    
    def _incremental_render(self, diff: Dict[str, Any], new_blocks: List[str]) -> None:
        """
        增量渲染。
        
        Args:
            diff: 差异信息
            new_blocks: 新块列表
        """
        # 对于简单实现，仍然使用全量渲染
        # 真正的增量渲染需要预览组件支持块级更新
        content = '\n\n'.join(new_blocks)
        self._full_render(content)
    
    def reset(self) -> None:
        """重置渲染器状态。"""
        self._last_content = ""
        self._last_blocks = []
        self._last_hashes = []
    
    def get_last_block_count(self) -> int:
        """获取上次渲染的块数量。"""
        return len(self._last_blocks)
    
    def force_full_render(self, content: str) -> None:
        """
        强制全量渲染。
        
        Args:
            content: Markdown 内容
        """
        self.reset()
        self.render(content)
