# -*- coding: utf-8 -*-
"""实时协作功能测试"""

import pytest
from unittest.mock import Mock, patch, MagicMock
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ui.features.collaboration.crdt import CRDTEngine, CRDTOperation
from ui.features.collaboration.cursor import CursorManager, RemoteCursor
from ui.features.collaboration.comments import CommentManager, Comment, CommentThread
from ui.features.collaboration.history import HistoryManager, HistoryEntry
from ui.features.collaboration.mentions import MentionManager, Mention, Task


class TestCRDTEngine:
    """CRDTEngine 测试"""
    
    def test_init(self):
        """测试初始化"""
        engine = CRDTEngine()
        assert engine.site_id is not None
        assert engine.get_content() == ""
    
    def test_set_content(self):
        """测试设置内容"""
        engine = CRDTEngine()
        engine.set_content("Hello World")
        
        assert engine.get_content() == "Hello World"
    
    def test_local_insert(self):
        """测试本地插入"""
        engine = CRDTEngine()
        engine.set_content("Hello")
        
        op = engine.local_insert(5, " World")
        
        assert engine.get_content() == "Hello World"
        assert op.type == 'insert'
        assert op.content == " World"
    
    def test_local_delete(self):
        """测试本地删除"""
        engine = CRDTEngine()
        engine.set_content("Hello World")
        
        op = engine.local_delete(5, 6)  # 删除 " World"
        
        assert engine.get_content() == "Hello"
        assert op.type == 'delete'
        assert op.length == 6
    
    def test_get_length(self):
        """测试获取长度"""
        engine = CRDTEngine()
        engine.set_content("Hello")
        
        assert engine.get_length() == 5
    
    def test_state_serialization(self):
        """测试状态序列化"""
        engine1 = CRDTEngine("site1")
        engine1.set_content("Hello World")
        
        state = engine1.get_state()
        
        engine2 = CRDTEngine("site2")
        engine2.load_state(state)
        
        assert engine2.get_content() == "Hello World"
    
    def test_operation_to_dict(self):
        """测试操作转字典"""
        op = CRDTOperation(
            id="test-id",
            type="insert",
            position=0,
            content="Hello",
            author="user1",
            timestamp=1234567890.0,
            vector_clock={"user1": 1}
        )
        
        d = op.to_dict()
        
        assert d['id'] == "test-id"
        assert d['type'] == "insert"
        assert d['content'] == "Hello"
    
    def test_operation_from_dict(self):
        """测试从字典创建操作"""
        d = {
            'id': "test-id",
            'type': "insert",
            'position': 0,
            'content': "Hello",
            'author': "user1",
            'timestamp': 1234567890.0,
            'vector_clock': {"user1": 1}
        }
        
        op = CRDTOperation.from_dict(d)
        
        assert op.id == "test-id"
        assert op.type == "insert"
        assert op.content == "Hello"


class TestCursorManager:
    """CursorManager 测试"""
    
    def test_add_cursor(self):
        """测试添加光标"""
        manager = CursorManager()
        
        color = manager.add_cursor("user1", "User 1")
        
        assert "user1" in manager.cursors
        assert manager.cursors["user1"].name == "User 1"
        assert color is not None
    
    def test_remove_cursor(self):
        """测试移除光标"""
        manager = CursorManager()
        manager.add_cursor("user1", "User 1")
        
        manager.remove_cursor("user1")
        
        assert "user1" not in manager.cursors
    
    def test_update_cursor(self):
        """测试更新光标"""
        manager = CursorManager()
        manager.add_cursor("user1", "User 1")
        
        manager.update_cursor("user1", 100, (50, 150))
        
        cursor = manager.cursors["user1"]
        assert cursor.position == 100
        assert cursor.selection_start == 50
        assert cursor.selection_end == 150
    
    def test_get_all_cursors(self):
        """测试获取所有光标"""
        manager = CursorManager()
        manager.add_cursor("user1", "User 1")
        manager.add_cursor("user2", "User 2")
        
        cursors = manager.get_all_cursors()
        
        assert len(cursors) == 2
    
    def test_clear_all(self):
        """测试清除所有光标"""
        manager = CursorManager()
        manager.add_cursor("user1", "User 1")
        manager.add_cursor("user2", "User 2")
        
        manager.clear_all()
        
        assert len(manager.cursors) == 0


class TestRemoteCursor:
    """RemoteCursor 测试"""
    
    def test_has_selection(self):
        """测试选区检测"""
        cursor_no_selection = RemoteCursor(
            participant_id="user1",
            name="User 1",
            color="#FF0000"
        )
        assert cursor_no_selection.has_selection is False
        
        cursor_with_selection = RemoteCursor(
            participant_id="user2",
            name="User 2",
            color="#00FF00",
            selection_start=10,
            selection_end=20
        )
        assert cursor_with_selection.has_selection is True


class TestCommentManager:
    """CommentManager 测试"""
    
    def test_create_thread(self):
        """测试创建评论线程"""
        manager = CommentManager()
        
        thread = manager.create_thread(
            start=0,
            end=10,
            initial_comment="This is a comment",
            author_id="user1",
            author_name="User 1"
        )
        
        assert thread.id is not None
        assert thread.document_range == (0, 10)
        assert len(thread.comments) == 1
    
    def test_add_reply(self):
        """测试添加回复"""
        manager = CommentManager()
        thread = manager.create_thread(0, 10, "Initial", "user1", "User 1")
        
        reply = manager.add_reply(thread.id, "Reply", "user2", "User 2")
        
        assert reply is not None
        assert len(manager.threads[thread.id].comments) == 2
    
    def test_resolve_thread(self):
        """测试解决线程"""
        manager = CommentManager()
        thread = manager.create_thread(0, 10, "Comment", "user1", "User 1")
        
        manager.resolve_thread(thread.id)
        
        assert manager.threads[thread.id].resolved is True
    
    def test_get_threads_in_range(self):
        """测试获取范围内的线程"""
        manager = CommentManager()
        manager.create_thread(0, 10, "Comment 1", "user1", "User 1")
        manager.create_thread(20, 30, "Comment 2", "user1", "User 1")
        manager.create_thread(50, 60, "Comment 3", "user1", "User 1")
        
        threads = manager.get_threads_in_range(5, 25)
        
        assert len(threads) == 2  # 第一个和第二个线程
    
    def test_export_comments(self):
        """测试导出评论"""
        manager = CommentManager()
        manager.create_thread(0, 10, "Test comment", "user1", "User 1")
        
        export = manager.export_comments()
        
        assert "# 评论" in export
        assert "Test comment" in export


class TestHistoryManager:
    """HistoryManager 测试"""
    
    def test_record(self):
        """测试记录历史"""
        manager = HistoryManager()
        
        entry = manager.record(
            author_id="user1",
            author_name="User 1",
            operation_type="edit",
            content_before="Hello",
            content_after="Hello World"
        )
        
        assert entry.id is not None
        assert len(manager.entries) == 1
    
    def test_get_history(self):
        """测试获取历史"""
        manager = HistoryManager()
        
        for i in range(5):
            manager.record("user1", "User 1", "edit", f"v{i}", f"v{i+1}")
        
        history = manager.get_history(limit=3)
        
        assert len(history) == 3
    
    def test_restore(self):
        """测试恢复版本"""
        manager = HistoryManager()
        entry = manager.record("user1", "User 1", "edit", "before", "after")
        
        content = manager.restore(entry.id)
        
        assert content == "after"
    
    def test_diff(self):
        """测试版本对比"""
        manager = HistoryManager()
        entry1 = manager.record("user1", "User 1", "edit", "", "Hello")
        entry2 = manager.record("user1", "User 1", "edit", "Hello", "Hello World")
        
        diff = manager.diff(entry1.id, entry2.id)
        
        assert isinstance(diff, list)
    
    def test_export_import(self):
        """测试导出导入"""
        manager1 = HistoryManager()
        manager1.record("user1", "User 1", "edit", "a", "b")
        
        data = manager1.export_history()
        
        manager2 = HistoryManager()
        manager2.import_history(data)
        
        assert len(manager2.entries) == 1


class TestMentionManager:
    """MentionManager 测试"""
    
    def test_parse_mentions(self):
        """测试解析提及"""
        manager = MentionManager()
        manager.set_participants({"alice": "user1", "bob": "user2"})
        
        content = "Hello @alice and @bob"
        mentions = manager.parse_mentions(content, "user3")
        
        assert len(mentions) == 2
    
    def test_parse_tasks(self):
        """测试解析任务"""
        manager = MentionManager()
        manager.set_participants({"alice": "user1"})
        
        content = "- [ ] Task 1\n- [x] Task 2 @alice\n- [ ] Task 3"
        tasks = manager.parse_tasks(content, "user2")
        
        assert len(tasks) == 3
        assert tasks[1].completed is True
    
    def test_complete_task(self):
        """测试完成任务"""
        manager = MentionManager()
        content = "- [ ] Task 1"
        tasks = manager.parse_tasks(content, "user1")
        
        task = manager.complete_task(tasks[0].id)
        
        assert task.completed is True
        assert task.completed_at is not None
    
    def test_get_user_suggestions(self):
        """测试获取用户建议"""
        manager = MentionManager()
        manager.set_participants({
            "alice": "user1",
            "alex": "user2",
            "bob": "user3"
        })
        
        suggestions = manager.get_user_suggestions("al")
        
        assert "alice" in suggestions
        assert "alex" in suggestions
        assert "bob" not in suggestions


# Property-based tests
try:
    from hypothesis import given, strategies as st, settings
    
    class TestCRDTProperties:
        """CRDT 属性测试"""
        
        @given(st.text(min_size=0, max_size=100))
        @settings(max_examples=20)
        def test_set_get_content_identity(self, content):
            """Property: set_content 后 get_content 应返回相同内容"""
            engine = CRDTEngine()
            engine.set_content(content)
            
            assert engine.get_content() == content
        
        @given(
            st.text(min_size=1, max_size=50),
            st.integers(min_value=0, max_value=50),
            st.text(min_size=1, max_size=10)
        )
        @settings(max_examples=20)
        def test_insert_increases_length(self, initial, pos, insert_text):
            """Property: 插入操作应增加文档长度"""
            engine = CRDTEngine()
            engine.set_content(initial)
            
            # 确保位置有效
            pos = min(pos, len(initial))
            
            initial_len = engine.get_length()
            engine.local_insert(pos, insert_text)
            
            assert engine.get_length() == initial_len + len(insert_text)

except ImportError:
    pass


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
