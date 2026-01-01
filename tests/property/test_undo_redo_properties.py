# -*- coding: utf-8 -*-
"""
撤销/重做属性测试
**Feature: markdown-editor-optimization, Property 1: 撤销/重做一致性**
**Validates: Requirements 2.1, 6.1**
"""

import pytest
from hypothesis import given, strategies as st, settings, assume
from typing import List
import sys
import os

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

from ui.features.undo_redo import UndoRedoManager, TextOperation


class MockTextWidget:
    """模拟Text组件用于测试 - 精确模拟位置操作"""
    
    def __init__(self):
        self.content = ""
        self._undo = True
    
    def _parse_index(self, index: str) -> int:
        """解析Text widget索引格式 (如 '1.5') 为字符位置"""
        if isinstance(index, int):
            return index
        index = str(index).strip()
        
        # 处理 "end" 和 "end-1c" 等特殊索引
        if index.startswith("end"):
            pos = len(self.content)
            if "-" in index:
                # 解析 "end-Nc" 格式
                try:
                    offset = int(index.split("-")[1].replace("c", "").replace("chars", ""))
                    pos = max(0, pos - offset)
                except:
                    pass
            return pos
        
        # 处理 "1.N + M chars" 格式 (先处理这个，因为它包含 ".")
        if "+" in index and "char" in index.lower():
            try:
                base, rest = index.split("+", 1)
                base_pos = self._parse_index(base.strip())
                # 提取数字
                offset_str = rest.strip().split()[0]
                offset = int(offset_str)
                return min(base_pos + offset, len(self.content))
            except Exception as e:
                return len(self.content)
        
        # 处理 "1.N" 格式 (行.列)
        if "." in index:
            try:
                parts = index.split(".")
                col = int(parts[1])
                return min(col, len(self.content))
            except:
                return 0
        
        return 0
    
    def get(self, start, end=None):
        start_pos = self._parse_index(start)
        if end is None:
            end_pos = start_pos + 1
        else:
            end_pos = self._parse_index(end)
        return self.content[start_pos:end_pos]
    
    def insert(self, index, chars, *args):
        """在指定位置插入文本"""
        pos = self._parse_index(index)
        self.content = self.content[:pos] + chars + self.content[pos:]
    
    def delete(self, index1, index2=None):
        """删除指定范围的文本"""
        start_pos = self._parse_index(index1)
        if index2 is None:
            end_pos = start_pos + 1
        else:
            end_pos = self._parse_index(index2)
        self.content = self.content[:start_pos] + self.content[end_pos:]
    
    def index(self, pos):
        return f"1.{self._parse_index(pos)}"
    
    def mark_set(self, mark, pos):
        pass
    
    def see(self, pos):
        pass
    
    def cget(self, option):
        if option == 'undo':
            return self._undo
        return None
    
    def configure(self, **kwargs):
        if 'undo' in kwargs:
            self._undo = kwargs['undo']


class TestUndoRedoProperties:
    """撤销/重做属性测试类"""
    
    @given(st.lists(st.text(alphabet=st.characters(whitelist_categories=('L', 'N')), 
                           min_size=1, max_size=5), 
                   min_size=1, max_size=10))
    @settings(max_examples=100, deadline=None)
    def test_undo_redo_roundtrip(self, operations: List[str]):
        """
        **Feature: markdown-editor-optimization, Property 1: 撤销/重做一致性**
        **Validates: Requirements 2.1, 6.1**
        
        对于任意编辑操作序列，撤销后重做应恢复原状态
        """
        # 过滤空字符串
        operations = [op for op in operations if op.strip()]
        assume(len(operations) > 0)
        
        widget = MockTextWidget()
        manager = UndoRedoManager(widget)
        
        # 执行所有插入操作 - 总是在末尾追加
        for op in operations:
            pos = f"1.{len(widget.content)}"
            widget.insert(pos, op)
            manager.record_insert(pos, op)
        
        content_before_undo = widget.content
        num_operations = len(operations)
        
        # 撤销所有操作
        undo_count = 0
        while manager.can_undo() and undo_count < num_operations + 5:  # 防止无限循环
            manager.undo()
            undo_count += 1
        
        content_after_undo = widget.content
        
        # 重做所有操作
        redo_count = 0
        while manager.can_redo() and redo_count < num_operations + 5:  # 防止无限循环
            manager.redo()
            redo_count += 1
        
        content_after_redo = widget.content
        
        # 验证：撤销和重做次数应该相等
        assert undo_count == redo_count, f"撤销次数({undo_count}) != 重做次数({redo_count})"
        
        # 验证：撤销后内容应该为空（因为我们撤销了所有操作）
        assert content_after_undo == "", f"撤销后内容应为空，实际: '{content_after_undo}'"
        
        # 验证：重做后内容应与撤销前一致
        assert content_before_undo == content_after_redo, \
            f"内容不一致: '{content_before_undo}' != '{content_after_redo}'"
    
    @given(st.integers(min_value=1, max_value=50))
    @settings(max_examples=50, deadline=None)
    def test_undo_stack_size_limit(self, num_operations: int):
        """
        测试撤销栈大小限制
        """
        widget = MockTextWidget()
        max_undo = 10
        manager = UndoRedoManager(widget, max_undo=max_undo)
        
        # 执行多次操作
        for i in range(num_operations):
            pos = f"1.{i}"
            manager.record_insert(pos, f"text{i}")
        
        # 验证：栈大小不超过限制
        assert len(manager.undo_stack) <= max_undo, \
            f"撤销栈大小({len(manager.undo_stack)}) > 限制({max_undo})"
    
    def test_undo_clears_redo_stack(self):
        """
        测试新操作清空重做栈
        """
        widget = MockTextWidget()
        manager = UndoRedoManager(widget)
        
        # 执行操作
        manager.record_insert("1.0", "hello")
        manager.record_insert("1.5", "world")
        
        # 撤销一次
        manager.undo()
        assert manager.can_redo()
        
        # 执行新操作
        manager.record_insert("1.5", "new")
        
        # 重做栈应该被清空
        assert not manager.can_redo(), "新操作后重做栈应该被清空"
    
    def test_manager_enabled_by_default(self):
        """
        测试管理器默认启用
        """
        widget = MockTextWidget()
        manager = UndoRedoManager(widget)
        
        assert manager.enabled, "UndoRedoManager应该默认启用"
    
    def test_disabled_manager_does_not_record(self):
        """
        测试禁用时不记录操作
        """
        widget = MockTextWidget()
        manager = UndoRedoManager(widget)
        
        # 禁用
        manager.disable()
        
        # 尝试记录
        manager.record_insert("1.0", "test")
        
        # 应该没有记录
        assert len(manager.undo_stack) == 0, "禁用时不应记录操作"
    
    def test_batch_operations_single_undo(self):
        """
        测试批量操作作为单个撤销点
        **Validates: Requirements 2.1**
        """
        widget = MockTextWidget()
        manager = UndoRedoManager(widget)
        
        # 开始批量操作
        manager.begin_batch()
        
        # 执行多个操作
        widget.insert("1.0", "line1\n")
        manager.record_insert("1.0", "line1\n")
        widget.insert("1.6", "line2\n")
        manager.record_insert("1.6", "line2\n")
        widget.insert("1.12", "line3")
        manager.record_insert("1.12", "line3")
        
        # 结束批量操作
        manager.end_batch()
        
        content_before = widget.content
        
        # 应该只有一个撤销点
        assert len(manager.undo_stack) == 1, f"批量操作应合并为1个撤销点，实际: {len(manager.undo_stack)}"
        
        # 一次撤销应该撤销所有批量操作
        manager.undo()
        
        assert widget.content == "", f"批量撤销后内容应为空，实际: '{widget.content}'"
        assert not manager.can_undo(), "批量撤销后不应有更多可撤销操作"
        
        # 重做应该恢复所有内容
        manager.redo()
        assert widget.content == content_before, f"批量重做后内容不一致"
    
    def test_batch_cancel_discards_operations(self):
        """
        测试取消批量操作不保存
        """
        widget = MockTextWidget()
        manager = UndoRedoManager(widget)
        
        # 开始批量操作
        manager.begin_batch()
        
        # 执行操作
        manager.record_insert("1.0", "test1")
        manager.record_insert("1.5", "test2")
        
        # 取消批量操作
        manager.cancel_batch()
        
        # 不应该有撤销点
        assert len(manager.undo_stack) == 0, "取消批量操作后不应有撤销点"
        assert not manager.in_batch_mode, "取消后不应处于批量模式"
    
    def test_empty_batch_not_recorded(self):
        """
        测试空批量操作不记录
        """
        widget = MockTextWidget()
        manager = UndoRedoManager(widget)
        
        # 开始并立即结束批量操作
        manager.begin_batch()
        manager.end_batch()
        
        # 不应该有撤销点
        assert len(manager.undo_stack) == 0, "空批量操作不应创建撤销点"
    
    @given(st.lists(st.text(alphabet=st.characters(whitelist_categories=('L', 'N')), 
                           min_size=1, max_size=5), 
                   min_size=2, max_size=8))
    @settings(max_examples=50, deadline=None)
    def test_batch_undo_redo_consistency(self, operations: List[str]):
        """
        **Property: 批量操作撤销/重做一致性**
        对于任意批量操作，撤销后重做应恢复原状态
        """
        operations = [op for op in operations if op.strip()]
        assume(len(operations) >= 2)
        
        widget = MockTextWidget()
        manager = UndoRedoManager(widget)
        
        # 开始批量操作
        manager.begin_batch()
        
        # 执行所有操作
        for op in operations:
            pos = f"1.{len(widget.content)}"
            widget.insert(pos, op)
            manager.record_insert(pos, op)
        
        # 结束批量操作
        manager.end_batch()
        
        content_before = widget.content
        
        # 撤销
        assert manager.can_undo()
        manager.undo()
        
        # 重做
        assert manager.can_redo()
        manager.redo()
        
        # 验证内容一致
        assert widget.content == content_before, \
            f"批量操作撤销/重做后内容不一致: '{widget.content}' != '{content_before}'"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
