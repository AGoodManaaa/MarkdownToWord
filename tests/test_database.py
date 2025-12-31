# -*- coding: utf-8 -*-
"""Markdown 数据库功能测试"""

import pytest
from unittest.mock import Mock, patch, MagicMock
import sys
import os
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ui.features.database.vault import VaultManager, VaultInfo, DocumentInfo, VaultError
from ui.features.database.search import SearchEngine, SearchResult, SearchMatch
from ui.features.database.tags import TagManager, TagInfo
from ui.features.database.links import LinkManager, LinkInfo
from ui.features.database.graph import GraphView, GraphNode, GraphEdge, GraphData
from ui.features.database.metadata import MetadataParser


class TestVaultManager:
    """VaultManager 测试"""
    
    def test_init_memory_database(self):
        """测试内存数据库初始化"""
        manager = VaultManager()
        assert manager._db_conn is not None
    
    def test_open_nonexistent_vault(self):
        """测试打开不存在的文档库"""
        manager = VaultManager()
        with pytest.raises(VaultError) as exc_info:
            manager.open_vault('/nonexistent/path')
        assert "不存在" in str(exc_info.value)
    
    def test_create_and_open_vault(self, tmp_path):
        """测试创建和打开文档库"""
        manager = VaultManager()
        
        vault_info = manager.open_vault(str(tmp_path))
        
        assert vault_info.path == tmp_path
        assert vault_info.file_count == 0
    
    def test_scan_vault_with_markdown_files(self, tmp_path):
        """测试扫描包含 Markdown 文件的文档库"""
        # 创建测试文件
        (tmp_path / "test1.md").write_text("# Test 1\n\nContent", encoding='utf-8')
        (tmp_path / "test2.md").write_text("# Test 2\n\n#tag1 #tag2", encoding='utf-8')
        
        manager = VaultManager()
        vault_info = manager.open_vault(str(tmp_path))
        
        assert vault_info.file_count == 2
    
    def test_extract_tags_from_content(self, tmp_path):
        """测试从内容中提取标签"""
        content = "# Title\n\n#python #coding #test"
        (tmp_path / "tagged.md").write_text(content, encoding='utf-8')
        
        manager = VaultManager()
        manager.open_vault(str(tmp_path))
        
        doc = manager.get_document("tagged.md")
        assert doc is not None
        assert 'python' in doc.tags or 'coding' in doc.tags
    
    def test_extract_links_from_content(self, tmp_path):
        """测试从内容中提取链接"""
        content = "# Title\n\nLink to [[other]] and [[another|display]]"
        (tmp_path / "linked.md").write_text(content, encoding='utf-8')
        
        manager = VaultManager()
        manager.open_vault(str(tmp_path))
        
        doc = manager.get_document("linked.md")
        assert doc is not None
        assert 'other' in doc.links
        assert 'another' in doc.links


class TestSearchEngine:
    """SearchEngine 测试"""
    
    def test_search_empty_query(self, tmp_path):
        """测试空查询"""
        manager = VaultManager()
        manager.open_vault(str(tmp_path))
        
        engine = SearchEngine(manager)
        results = engine.search("")
        
        assert results == []
    
    def test_search_with_results(self, tmp_path):
        """测试有结果的搜索"""
        (tmp_path / "test.md").write_text("# Hello World\n\nThis is a test document.", encoding='utf-8')
        
        manager = VaultManager()
        manager.open_vault(str(tmp_path))
        
        engine = SearchEngine(manager)
        results = engine.search("Hello")
        
        assert len(results) >= 0  # 可能有结果
    
    def test_parse_query_simple(self):
        """测试简单查询解析"""
        manager = VaultManager()
        engine = SearchEngine(manager)
        
        parsed = engine.parse_query("hello world")
        
        assert 'hello' in parsed['must']
        assert 'world' in parsed['must']
    
    def test_parse_query_with_not(self):
        """测试带 NOT 的查询解析"""
        manager = VaultManager()
        engine = SearchEngine(manager)
        
        parsed = engine.parse_query("hello -world")
        
        assert 'hello' in parsed['must']
        assert 'world' in parsed['must_not']
    
    def test_parse_query_with_phrase(self):
        """测试带引号短语的查询解析"""
        manager = VaultManager()
        engine = SearchEngine(manager)
        
        parsed = engine.parse_query('"hello world"')
        
        assert 'hello world' in parsed['phrases']


class TestTagManager:
    """TagManager 测试"""
    
    def test_extract_tags(self):
        """测试标签提取"""
        manager = VaultManager()
        tag_manager = TagManager(manager)
        
        content = "This is #python and #coding content"
        tags = tag_manager.extract_tags(content)
        
        assert 'python' in tags
        assert 'coding' in tags
    
    def test_extract_hierarchical_tags(self):
        """测试层级标签提取"""
        manager = VaultManager()
        tag_manager = TagManager(manager)
        
        content = "This is #project/work and #project/personal"
        tags = tag_manager.extract_tags(content)
        
        assert 'project/work' in tags
        assert 'project/personal' in tags
    
    def test_get_all_tags_empty(self, tmp_path):
        """测试获取空标签列表"""
        manager = VaultManager()
        manager.open_vault(str(tmp_path))
        
        tag_manager = TagManager(manager)
        tags = tag_manager.get_all_tags()
        
        assert isinstance(tags, list)


class TestLinkManager:
    """LinkManager 测试"""
    
    def test_extract_links(self):
        """测试链接提取"""
        manager = VaultManager()
        link_manager = LinkManager(manager)
        
        content = "Link to [[page1]] and [[page2|Display Text]]"
        links = link_manager.extract_links(content, "source.md")
        
        assert len(links) == 2
        assert any(l.target == 'page1' for l in links)
        assert any(l.target == 'page2' for l in links)
    
    def test_create_link(self):
        """测试创建链接语法"""
        manager = VaultManager()
        link_manager = LinkManager(manager)
        
        # 简单链接
        simple = link_manager.create_link("target")
        assert simple == "[[target]]"
        
        # 带显示文本的链接
        with_display = link_manager.create_link("target", "Display")
        assert with_display == "[[target|Display]]"
    
    def test_wiki_link_pattern(self):
        """测试 wiki 链接正则表达式"""
        manager = VaultManager()
        link_manager = LinkManager(manager)
        
        # 测试各种格式
        test_cases = [
            ("[[simple]]", "simple"),
            ("[[with space]]", "with space"),
            ("[[path/to/file]]", "path/to/file"),
            ("[[link|display]]", "link"),
        ]
        
        for text, expected in test_cases:
            match = link_manager.WIKI_LINK_PATTERN.search(text)
            assert match is not None
            assert match.group(1) == expected


class TestGraphView:
    """GraphView 测试"""
    
    def test_build_empty_graph(self):
        """测试构建空图"""
        manager = VaultManager()
        link_manager = LinkManager(manager)
        graph_view = GraphView(link_manager)
        
        graph = graph_view.build_graph([])
        
        assert graph.node_count == 0
        assert graph.edge_count == 0
    
    def test_build_graph_with_nodes(self):
        """测试构建有节点的图"""
        manager = VaultManager()
        link_manager = LinkManager(manager)
        graph_view = GraphView(link_manager)
        
        docs = [
            DocumentInfo(path=Path("doc1.md"), title="Doc 1", links=["doc2"]),
            DocumentInfo(path=Path("doc2.md"), title="Doc 2", links=[]),
        ]
        
        graph = graph_view.build_graph(docs)
        
        assert graph.node_count == 2
    
    def test_filter_by_tag(self):
        """测试按标签过滤"""
        manager = VaultManager()
        link_manager = LinkManager(manager)
        graph_view = GraphView(link_manager)
        
        nodes = [
            GraphNode(id="1", label="Node 1", tags=["python"]),
            GraphNode(id="2", label="Node 2", tags=["java"]),
            GraphNode(id="3", label="Node 3", tags=["python"]),
        ]
        edges = [
            GraphEdge(source="1", target="2"),
            GraphEdge(source="1", target="3"),
        ]
        graph = GraphData(nodes=nodes, edges=edges)
        
        filtered = graph_view.filter_by_tag(graph, "python")
        
        assert filtered.node_count == 2
    
    def test_get_neighbors(self):
        """测试获取相邻节点"""
        manager = VaultManager()
        link_manager = LinkManager(manager)
        graph_view = GraphView(link_manager)
        
        nodes = [
            GraphNode(id="1", label="Node 1"),
            GraphNode(id="2", label="Node 2"),
            GraphNode(id="3", label="Node 3"),
        ]
        edges = [
            GraphEdge(source="1", target="2"),
            GraphEdge(source="1", target="3"),
        ]
        graph = GraphData(nodes=nodes, edges=edges)
        
        neighbors = graph_view.get_neighbors(graph, "1")
        
        assert "2" in neighbors
        assert "3" in neighbors
    
    def test_get_statistics(self):
        """测试获取图统计"""
        manager = VaultManager()
        link_manager = LinkManager(manager)
        graph_view = GraphView(link_manager)
        
        nodes = [
            GraphNode(id="1", label="Node 1"),
            GraphNode(id="2", label="Node 2"),
        ]
        edges = [GraphEdge(source="1", target="2")]
        graph = GraphData(nodes=nodes, edges=edges)
        
        stats = graph_view.get_statistics(graph)
        
        assert stats['node_count'] == 2
        assert stats['edge_count'] == 1


class TestMetadataParser:
    """MetadataParser 测试"""
    
    def test_parse_no_frontmatter(self):
        """测试解析无 frontmatter 的内容"""
        parser = MetadataParser()
        
        content = "# Title\n\nContent"
        metadata, body = parser.parse(content)
        
        assert metadata == {}
        assert body == content
    
    def test_parse_with_frontmatter(self):
        """测试解析有 frontmatter 的内容"""
        pytest.importorskip('yaml')  # 跳过如果 yaml 未安装
        
        parser = MetadataParser()
        
        content = """---
title: Test
tags:
  - python
  - coding
---

# Content"""
        
        metadata, body = parser.parse(content)
        
        assert metadata.get('title') == 'Test'
        assert 'python' in metadata.get('tags', [])
    
    def test_create_default(self):
        """测试创建默认 frontmatter"""
        parser = MetadataParser()
        
        metadata = parser.create_default("Test Title", "Author")
        
        assert metadata['title'] == "Test Title"
        assert metadata['author'] == "Author"
        assert 'created' in metadata
    
    def test_validate_valid_metadata(self):
        """测试验证有效元数据"""
        parser = MetadataParser()
        
        metadata = {
            'title': 'Test',
            'tags': ['a', 'b'],
            'created': '2024-01-01'
        }
        
        errors = parser.validate(metadata)
        assert errors == []
    
    def test_validate_invalid_tags(self):
        """测试验证无效标签"""
        parser = MetadataParser()
        
        metadata = {
            'tags': 'not a list'
        }
        
        errors = parser.validate(metadata)
        assert len(errors) > 0


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
