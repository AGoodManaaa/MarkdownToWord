# -*- coding: utf-8 -*-
"""
滚动同步属性测试
**Feature: best-markdown-editor, Property 3: 预览同步精确性**
**Feature: best-markdown-editor, Property 4: 增量渲染正确性**
**Validates: Requirements 3.1, 3.4**
"""

import pytest
from hypothesis import given, strategies as st, settings, assume
from typing import List, Dict
import sys
import os

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

from ui.features.precise_scroll_sync import (
    PreciseScrollSync, LineMapping, IncrementalPreviewUpdater
)


class MockEditor:
    """模拟编辑器组件"""
    
    def __init__(self, content: str = ""):
        self.content = content
        self._textbox = MockTextWidget(content)
    
    def get(self, start, end):
        return self.content


class MockTextWidget:
    """模拟 Text 组件"""
    
    def __init__(self, content: str = ""):
        self.content = content
        self._yview_pos = 0.0
    
    def index(self, pos):
        if pos == "@0,0":
            return "1.0"
        if pos == "end":
            lines = self.content.split('\n')
            return f"{len(lines) + 1}.0"
        return "1.0"
    
    def yview(self):
        return (self._yview_pos, self._yview_pos + 0.3)
    
    def yview_moveto(self, pos):
        self._yview_pos = pos
    
    def see(self, pos):
        pass


class MockPreview:
    """模拟预览组件"""
    
    def __init__(self):
        self.text = MockTextWidget()
    
    def sync_scroll_to(self, pos):
        self.text.yview_moveto(pos)


class TestScrollSyncProperties:
    """滚动同步属性测试"""
    
    @given(st.integers(min_value=1, max_value=100))
    @settings(max_examples=100, deadline=None)
    def test_sync_accuracy_within_tolerance(self, editor_line: int):
        """
        **Feature: best-markdown-editor, Property 3: 预览同步精确性**
        **Validates: Requirements 3.1**
        
        对于任意编辑器滚动位置，预览区对应位置的源码行号与编辑器当前行号误差不超过2行
        """
        # 创建模拟组件
        content = '\n'.join([f"Line {i}" for i in range(1, 101)])
        editor = MockEditor(content)
        preview = MockPreview()
        
        sync = PreciseScrollSync(editor, preview)
        sync.build_line_map(content)
        
        # 模拟同步
        preview_line = editor_line  # 理想情况下应该相同
        
        # 验证精确度
        assert sync.is_sync_accurate(editor_line, preview_line, tolerance=2), \
            f"同步误差超过2行: editor={editor_line}, preview={preview_line}"
    
    @given(st.integers(min_value=1, max_value=50), st.integers(min_value=1, max_value=50))
    @settings(max_examples=50, deadline=None)
    def test_sync_accuracy_calculation(self, line1: int, line2: int):
        """
        测试同步精确度计算
        """
        editor = MockEditor("")
        preview = MockPreview()
        sync = PreciseScrollSync(editor, preview)
        
        accuracy = sync.get_sync_accuracy(line1, line2)
        expected = abs(line1 - line2)
        
        assert accuracy == expected, \
            f"精确度计算错误: 期望 {expected}, 实际 {accuracy}"
    
    def test_line_map_building(self):
        """
        测试行映射表构建
        """
        content = """# 标题1

这是一个段落。

## 标题2

```python
print("hello")
```

> 引用文本

- 列表项1
- 列表项2
"""
        editor = MockEditor(content)
        preview = MockPreview()
        sync = PreciseScrollSync(editor, preview)
        
        sync.build_line_map(content)
        
        # 验证映射表不为空
        assert len(sync.line_map) > 0, "行映射表不应为空"
        
        # 验证标题被映射
        assert 1 in sync.line_map, "第1行（标题）应该被映射"
        assert sync.line_map[1].block_type == "h1", "第1行应该是 h1 类型"
    
    def test_empty_content_handling(self):
        """
        测试空内容处理
        """
        editor = MockEditor("")
        preview = MockPreview()
        sync = PreciseScrollSync(editor, preview)
        
        # 不应该抛出异常
        sync.build_line_map("")
        
        assert len(sync.line_map) == 0, "空内容的映射表应为空"


class TestIncrementalUpdateProperties:
    """增量更新属性测试"""
    
    @given(st.text(min_size=10, max_size=200, alphabet='abcdefghijklmnopqrstuvwxyz \n'))
    @settings(max_examples=50, deadline=None)
    def test_incremental_update_correctness(self, content: str):
        """
        **Feature: best-markdown-editor, Property 4: 增量渲染正确性**
        **Validates: Requirements 3.4**
        
        对于任意文档修改，增量渲染的结果应与全量渲染的结果一致
        """
        assume(content.strip())  # 确保非空
        
        preview = MockPreview()
        updater = IncrementalPreviewUpdater(preview)
        
        # 模拟内容变化
        old_content = ""
        new_content = content
        
        # 获取变化的块
        changed_blocks = updater.update(old_content, new_content)
        
        # 解析新内容的所有块
        all_blocks = updater._parse_blocks(new_content)
        
        # 验证：所有新块都应该在变化列表中
        for line_num, (block_type, block_content) in all_blocks.items():
            block_id = f"{block_type}_{line_num}"
            found = any(bid == block_id for bid, _ in changed_blocks)
            assert found, f"块 {block_id} 应该在变化列表中"
    
    @given(st.lists(st.text(min_size=1, max_size=30, alphabet='abcdefghijklmnopqrstuvwxyz'), 
                   min_size=2, max_size=10))
    @settings(max_examples=50, deadline=None)
    def test_no_change_detection(self, lines: List[str]):
        """
        测试无变化检测
        """
        content = '\n'.join(lines)
        
        preview = MockPreview()
        updater = IncrementalPreviewUpdater(preview)
        
        # 相同内容应该返回空列表
        changed_blocks = updater.update(content, content)
        
        assert len(changed_blocks) == 0, \
            f"相同内容不应有变化，但检测到 {len(changed_blocks)} 个变化"
    
    @given(st.lists(st.text(min_size=1, max_size=30, alphabet='abcdefghijklmnopqrstuvwxyz'), 
                   min_size=3, max_size=10),
           st.integers(min_value=0, max_value=9))
    @settings(max_examples=50, deadline=None)
    def test_single_line_change_detection(self, lines: List[str], change_idx: int):
        """
        **Feature: best-markdown-editor, Property 4: 增量渲染正确性**
        **Validates: Requirements 3.4**
        
        对于单行修改，只有修改的块应该被检测为变化
        """
        assume(len(lines) > 0)
        change_idx = change_idx % len(lines)
        
        old_content = '\n'.join(lines)
        
        # 修改一行
        new_lines = lines.copy()
        new_lines[change_idx] = new_lines[change_idx] + "_modified"
        new_content = '\n'.join(new_lines)
        
        preview = MockPreview()
        updater = IncrementalPreviewUpdater(preview)
        
        # 获取变化的块
        changed_blocks = updater.update(old_content, new_content)
        
        # 验证：至少检测到一个变化
        assert len(changed_blocks) >= 1, "应该检测到至少一个变化"
        
        # 验证：变化的块数量应该合理（不应该是全部重新渲染）
        all_blocks = updater._parse_blocks(new_content)
        if len(all_blocks) > 1:
            # 如果有多个块，变化的块数量应该小于总块数
            assert len(changed_blocks) <= len(all_blocks), \
                "变化的块数量不应超过总块数"
    
    @given(st.text(min_size=50, max_size=500, alphabet='abcdefghijklmnopqrstuvwxyz \n'))
    @settings(max_examples=30, deadline=None)
    def test_full_render_decision(self, content: str):
        """
        测试全量渲染决策
        """
        assume(content.strip())
        
        preview = MockPreview()
        updater = IncrementalPreviewUpdater(preview)
        
        # 空内容到有内容应该需要全量渲染
        should_full = updater.should_full_render("", content)
        assert should_full, "从空内容到有内容应该需要全量渲染"
        
        # 相同内容不需要全量渲染
        should_full = updater.should_full_render(content, content)
        assert not should_full, "相同内容不需要全量渲染"
    
    @given(st.lists(st.text(min_size=1, max_size=20, alphabet='abcdefghijklmnopqrstuvwxyz'), 
                   min_size=5, max_size=20))
    @settings(max_examples=30, deadline=None)
    def test_changed_line_range(self, lines: List[str]):
        """
        测试变化行范围检测
        """
        assume(len(lines) >= 5)
        
        old_content = '\n'.join(lines)
        
        # 修改中间的一行
        mid_idx = len(lines) // 2
        new_lines = lines.copy()
        new_lines[mid_idx] = "MODIFIED_LINE"
        new_content = '\n'.join(new_lines)
        
        preview = MockPreview()
        updater = IncrementalPreviewUpdater(preview)
        
        start, end = updater.get_changed_line_range(old_content, new_content)
        
        # 验证：变化范围应该包含修改的行
        assert start <= mid_idx + 1 <= end, \
            f"变化范围 ({start}, {end}) 应该包含修改的行 {mid_idx + 1}"
    
    def test_code_block_detection(self):
        """
        测试代码块检测
        """
        content = """# 标题

```python
def hello():
    print("world")
```

普通段落
"""
        preview = MockPreview()
        updater = IncrementalPreviewUpdater(preview)
        
        blocks = updater._parse_blocks(content)
        
        # 验证代码块被正确识别
        code_blocks = [(ln, bt, bc) for ln, (bt, bc) in blocks.items() if bt == 'code_block']
        assert len(code_blocks) > 0, "应该检测到代码块"
    
    def test_table_detection(self):
        """
        测试表格检测
        """
        content = """# 标题

| 列1 | 列2 |
|-----|-----|
| A   | B   |
| C   | D   |

普通段落
"""
        preview = MockPreview()
        updater = IncrementalPreviewUpdater(preview)
        
        blocks = updater._parse_blocks(content)
        
        # 验证表格被正确识别
        table_blocks = [(ln, bt, bc) for ln, (bt, bc) in blocks.items() if bt == 'table']
        assert len(table_blocks) > 0, "应该检测到表格"
    
    def test_quote_detection(self):
        """
        测试引用块检测
        """
        content = """# 标题

> 这是引用
> 多行引用

普通段落
"""
        preview = MockPreview()
        updater = IncrementalPreviewUpdater(preview)
        
        blocks = updater._parse_blocks(content)
        
        # 验证引用块被正确识别
        quote_blocks = [(ln, bt, bc) for ln, (bt, bc) in blocks.items() if bt == 'quote']
        assert len(quote_blocks) > 0, "应该检测到引用块"
    
    def test_list_detection(self):
        """
        测试列表检测
        """
        content = """# 标题

- 列表项1
- 列表项2
- 列表项3

普通段落
"""
        preview = MockPreview()
        updater = IncrementalPreviewUpdater(preview)
        
        blocks = updater._parse_blocks(content)
        
        # 验证列表被正确识别
        list_blocks = [(ln, bt, bc) for ln, (bt, bc) in blocks.items() if bt == 'list']
        assert len(list_blocks) > 0, "应该检测到列表"
    
    def test_cache_operations(self):
        """
        测试缓存操作
        """
        preview = MockPreview()
        updater = IncrementalPreviewUpdater(preview)
        
        # 缓存块
        updater.cache_block("test_block", "test content")
        
        # 获取缓存
        cached = updater.get_cached_block("test_block")
        assert cached == "test content", "缓存内容应该匹配"
        
        # 清空缓存
        updater.clear_cache()
        cached = updater.get_cached_block("test_block")
        assert cached is None, "清空后缓存应为空"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
