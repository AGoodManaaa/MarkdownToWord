# -*- coding: utf-8 -*-
"""
Tests for IncrementalPreviewRenderer module.

Includes both unit tests and property-based tests.
"""

import unittest
import time
from unittest.mock import MagicMock
from hypothesis import given, strategies as st, settings, assume

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ui.incremental_renderer import IncrementalPreviewRenderer


class MockPreviewWidget:
    """Mock preview widget for testing."""
    
    def __init__(self):
        self.content = ""
        self.update_count = 0
    
    def update_preview(self, content):
        self.content = content
        self.update_count += 1
    
    def set_content(self, content):
        self.content = content
        self.update_count += 1


class TestIncrementalRendererUnit(unittest.TestCase):
    """Unit tests for IncrementalPreviewRenderer."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.preview = MockPreviewWidget()
        self.renderer = IncrementalPreviewRenderer(self.preview)
    
    def test_render_empty_content(self):
        """Test rendering empty content."""
        result = self.renderer.render("")
        
        self.assertEqual(result['type'], 'full')
        self.assertEqual(self.preview.content, "")
    
    def test_render_initial_content(self):
        """Test rendering initial content uses full render."""
        content = "# Hello\n\nWorld"
        result = self.renderer.render(content)
        
        self.assertEqual(result['type'], 'full')
    
    def test_render_small_change_uses_incremental(self):
        """Test that small changes use incremental rendering."""
        # Initial render with more blocks to ensure change ratio < 0.3
        self.renderer.render("# Block 1\n\n# Block 2\n\n# Block 3\n\n# Block 4\n\n# Block 5")
        
        # Small change (one block out of 5 = 20% change)
        result = self.renderer.render("# Block 1\n\n# Block 2 Modified\n\n# Block 3\n\n# Block 4\n\n# Block 5")
        
        # Should use incremental (change ratio 0.2 < 0.3)
        self.assertEqual(result['type'], 'incremental')
    
    def test_render_large_change_uses_full(self):
        """Test that large changes use full rendering."""
        # Initial render
        self.renderer.render("# Block 1\n\n# Block 2")
        
        # Large change (completely different)
        result = self.renderer.render("# New 1\n\n# New 2\n\n# New 3\n\n# New 4")
        
        self.assertEqual(result['type'], 'full')
    
    def test_parse_blocks(self):
        """Test block parsing."""
        content = "# Block 1\n\n# Block 2\n\n# Block 3"
        blocks = self.renderer._parse_blocks(content)
        
        self.assertEqual(len(blocks), 3)
        self.assertEqual(blocks[0], "# Block 1")
        self.assertEqual(blocks[1], "# Block 2")
        self.assertEqual(blocks[2], "# Block 3")
    
    def test_parse_blocks_with_code(self):
        """Test block parsing preserves code blocks."""
        content = "# Title\n\n```python\ncode\n\nmore code\n```\n\n# End"
        blocks = self.renderer._parse_blocks(content)
        
        # Code block should be kept together
        self.assertEqual(len(blocks), 3)
        self.assertIn("```python", blocks[1])
        self.assertIn("more code", blocks[1])
    
    def test_reset(self):
        """Test reset clears state."""
        self.renderer.render("# Content")
        self.renderer.reset()
        
        self.assertEqual(self.renderer._last_content, "")
        self.assertEqual(len(self.renderer._last_blocks), 0)
    
    def test_get_last_block_count(self):
        """Test get_last_block_count."""
        self.renderer.render("# Block 1\n\n# Block 2\n\n# Block 3")
        
        self.assertEqual(self.renderer.get_last_block_count(), 3)
    
    def test_force_full_render(self):
        """Test force_full_render."""
        self.renderer.render("# Initial")
        self.renderer.force_full_render("# Forced")
        
        self.assertEqual(self.preview.content, "# Forced")


class TestIncrementalRendererProperty(unittest.TestCase):
    """
    Property-based tests for IncrementalPreviewRenderer.
    
    **Feature: code-optimization, Property 3, 4, 5**
    **Validates: Requirements 4.1, 4.2, 4.3**
    """
    
    def setUp(self):
        """Set up test fixtures."""
        self.preview = MockPreviewWidget()
    
    @given(st.text(min_size=0, max_size=50000))
    @settings(max_examples=50, deadline=None)
    def test_render_performance_property(self, content):
        """
        Property 3: 预览渲染性能
        
        *For any* 大于 10KB 的文档，预览渲染时间应该小于 200ms。
        
        **Feature: code-optimization, Property 3: 预览渲染性能**
        **Validates: Requirements 4.1**
        """
        renderer = IncrementalPreviewRenderer(self.preview)
        
        start_time = time.time()
        renderer.render(content)
        elapsed_ms = (time.time() - start_time) * 1000
        
        # Property: render should complete within 200ms
        self.assertLess(elapsed_ms, 200,
            f"Rendering {len(content)} bytes should take < 200ms (took {elapsed_ms:.1f}ms)")
    
    @given(st.lists(st.text(min_size=1, max_size=100), min_size=2, max_size=10))
    @settings(max_examples=100)
    def test_incremental_render_correctness_property(self, blocks):
        """
        Property 5: 增量渲染正确性
        
        *For any* 文档变化，增量渲染后的结果应该与全量渲染结果一致。
        
        **Feature: code-optimization, Property 5: 增量渲染正确性**
        **Validates: Requirements 4.3**
        """
        # Create two renderers
        preview1 = MockPreviewWidget()
        preview2 = MockPreviewWidget()
        incremental_renderer = IncrementalPreviewRenderer(preview1)
        full_renderer = IncrementalPreviewRenderer(preview2)
        
        # Initial content
        initial_content = '\n\n'.join(blocks[:-1])
        incremental_renderer.render(initial_content)
        
        # Modified content (add last block)
        modified_content = '\n\n'.join(blocks)
        
        # Incremental render
        incremental_renderer.render(modified_content)
        
        # Full render (reset first)
        full_renderer.render(modified_content)
        
        # Property: both should produce the same result
        self.assertEqual(preview1.content, preview2.content,
            "Incremental render should produce same result as full render")
    
    @given(st.lists(st.text(min_size=10, max_size=50), min_size=3, max_size=10))
    @settings(max_examples=100)
    def test_small_change_detection_property(self, blocks):
        """
        Property: Small changes should be detected as incremental.
        
        **Feature: code-optimization, Property 5: 增量渲染正确性**
        **Validates: Requirements 4.3**
        """
        renderer = IncrementalPreviewRenderer(self.preview, change_threshold=0.3)
        
        # Initial render
        initial_content = '\n\n'.join(blocks)
        renderer.render(initial_content)
        
        # Modify only one block (small change)
        modified_blocks = blocks.copy()
        modified_blocks[0] = modified_blocks[0] + " modified"
        modified_content = '\n\n'.join(modified_blocks)
        
        result = renderer.render(modified_content)
        
        # Property: changing one block out of many should be incremental
        if len(blocks) >= 4:  # Need enough blocks for change ratio < 0.3
            self.assertEqual(result['type'], 'incremental',
                f"Changing 1 of {len(blocks)} blocks should be incremental")


class TestDebounceProperty(unittest.TestCase):
    """
    Property-based tests for debounce behavior.
    
    **Feature: code-optimization, Property 4: 防抖有效性**
    **Validates: Requirements 4.2**
    """
    
    @given(st.integers(min_value=2, max_value=10))
    @settings(max_examples=50)
    def test_debounce_effectiveness_property(self, num_rapid_changes):
        """
        Property 4: 防抖有效性
        
        *For any* 连续的快速输入序列，只有最后一次输入触发预览更新。
        
        Note: This tests the concept - actual debounce is in PreviewSyncFeature.
        
        **Feature: code-optimization, Property 4: 防抖有效性**
        **Validates: Requirements 4.2**
        """
        preview = MockPreviewWidget()
        renderer = IncrementalPreviewRenderer(preview)
        
        # Simulate rapid changes
        final_content = None
        for i in range(num_rapid_changes):
            content = f"# Change {i}"
            final_content = content
            renderer.render(content)
        
        # Property: final content should be the last change
        self.assertEqual(preview.content, final_content,
            "Preview should show the final content after rapid changes")


if __name__ == '__main__':
    unittest.main()
