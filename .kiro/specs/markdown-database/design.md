# Design Document: Markdown 数据库

## Overview

本设计文档描述了 Markdown 数据库功能的技术实现方案。该功能将应用转变为知识管理系统，支持文档库管理、全文搜索、标签分类、双向链接和知识图谱可视化。

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                    Markdown Database                             │
├─────────────────────────────────────────────────────────────────┤
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐          │
│  │ VaultManager │  │ SearchEngine │  │ GraphView    │          │
│  └──────────────┘  └──────────────┘  └──────────────┘          │
│         │                │                  │                   │
│         ▼                ▼                  ▼                   │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐          │
│  │ FileWatcher  │  │ IndexManager │  │ LinkResolver │          │
│  │ MetadataDB   │  │ QueryParser  │  │ NodeLayout   │          │
│  └──────────────┘  └──────────────┘  └──────────────┘          │
│                          │                                      │
│                          ▼                                      │
│                   ┌──────────────┐                              │
│                   │   SQLite DB  │                              │
│                   └──────────────┘                              │
└─────────────────────────────────────────────────────────────────┘
```

## Components and Interfaces

### 1. VaultManager - 文档库管理器

```python
from pathlib import Path
from typing import List, Optional, Callable
from dataclasses import dataclass
from datetime import datetime

@dataclass
class VaultInfo:
    """文档库信息"""
    path: Path
    name: str
    file_count: int
    last_indexed: datetime
    size_bytes: int

@dataclass
class DocumentInfo:
    """文档信息"""
    path: Path
    title: str
    created: datetime
    modified: datetime
    size: int
    tags: List[str]
    links: List[str]
    backlinks: List[str]
    frontmatter: dict

class VaultManager:
    """文档库管理器"""
    
    def __init__(self, db_path: str = None):
        self.db_path = db_path
        self.current_vault: Optional[Path] = None
        self._file_watcher = None
        self._on_file_change: Optional[Callable] = None
    
    def open_vault(self, path: str) -> VaultInfo:
        """打开文档库"""
        pass
    
    def create_vault(self, path: str, name: str) -> VaultInfo:
        """创建新文档库"""
        pass
    
    def scan_vault(self) -> List[DocumentInfo]:
        """扫描文档库中的所有文件"""
        pass
    
    def watch_changes(self, callback: Callable[[str, str], None]) -> None:
        """监控文件变化 (path, event_type)"""
        pass
    
    def stop_watching(self) -> None:
        """停止监控"""
        pass
    
    def get_document(self, path: str) -> Optional[DocumentInfo]:
        """获取文档信息"""
        pass
    
    def get_all_documents(self) -> List[DocumentInfo]:
        """获取所有文档"""
        pass
    
    def get_recent_documents(self, limit: int = 10) -> List[DocumentInfo]:
        """获取最近修改的文档"""
        pass
```

### 2. SearchEngine - 全文搜索引擎

```python
from dataclasses import dataclass
from typing import List, Optional
from enum import Enum

class SearchOperator(Enum):
    AND = "AND"
    OR = "OR"
    NOT = "NOT"
    PHRASE = "PHRASE"

@dataclass
class SearchResult:
    """搜索结果"""
    document: DocumentInfo
    score: float
    matches: List['SearchMatch']
    
@dataclass
class SearchMatch:
    """匹配详情"""
    line_number: int
    content: str
    highlight_ranges: List[tuple]  # (start, end)

class SearchEngine:
    """全文搜索引擎"""
    
    def __init__(self, index_path: str):
        self.index_path = index_path
        self._index = None
    
    def build_index(self, documents: List[DocumentInfo]) -> None:
        """构建搜索索引"""
        pass
    
    def update_index(self, document: DocumentInfo) -> None:
        """更新单个文档的索引"""
        pass
    
    def remove_from_index(self, path: str) -> None:
        """从索引中移除文档"""
        pass
    
    def search(self, query: str, limit: int = 50) -> List[SearchResult]:
        """执行搜索"""
        pass
    
    def search_by_tag(self, tag: str) -> List[DocumentInfo]:
        """按标签搜索"""
        pass
    
    def search_by_filename(self, pattern: str) -> List[DocumentInfo]:
        """按文件名搜索"""
        pass
    
    def parse_query(self, query: str) -> dict:
        """解析高级查询语法"""
        pass
    
    def get_suggestions(self, prefix: str) -> List[str]:
        """获取搜索建议"""
        pass
```

### 3. TagManager - 标签管理器

```python
@dataclass
class TagInfo:
    """标签信息"""
    name: str
    count: int
    parent: Optional[str]  # 层级标签的父标签
    children: List[str]

class TagManager:
    """标签管理器"""
    
    def __init__(self, db_connection):
        self.db = db_connection
    
    def extract_tags(self, content: str) -> List[str]:
        """从文档内容提取标签"""
        pass
    
    def get_all_tags(self) -> List[TagInfo]:
        """获取所有标签"""
        pass
    
    def get_documents_by_tag(self, tag: str) -> List[DocumentInfo]:
        """获取包含指定标签的文档"""
        pass
    
    def rename_tag(self, old_name: str, new_name: str) -> int:
        """重命名标签，返回受影响的文档数"""
        pass
    
    def delete_tag(self, tag: str) -> int:
        """删除标签，返回受影响的文档数"""
        pass
    
    def get_tag_hierarchy(self) -> dict:
        """获取标签层级结构"""
        pass
    
    def merge_tags(self, source: str, target: str) -> int:
        """合并标签"""
        pass
```

### 4. LinkManager - 链接管理器

```python
import re
from typing import Set

@dataclass
class LinkInfo:
    """链接信息"""
    source: str  # 源文档路径
    target: str  # 目标文档路径
    link_text: str  # 链接文本
    line_number: int
    is_broken: bool

class LinkManager:
    """双向链接管理器"""
    
    WIKI_LINK_PATTERN = re.compile(r'\[\[([^\]|]+)(?:\|([^\]]+))?\]\]')
    
    def __init__(self, vault_manager: VaultManager):
        self.vault = vault_manager
        self._link_cache: dict = {}
    
    def extract_links(self, content: str, source_path: str) -> List[LinkInfo]:
        """从文档内容提取链接"""
        pass
    
    def get_outgoing_links(self, document_path: str) -> List[LinkInfo]:
        """获取文档的出链"""
        pass
    
    def get_backlinks(self, document_path: str) -> List[LinkInfo]:
        """获取文档的反向链接"""
        pass
    
    def resolve_link(self, link_text: str, from_path: str) -> Optional[str]:
        """解析链接目标路径"""
        pass
    
    def update_links_on_rename(self, old_path: str, new_path: str) -> int:
        """文档重命名时更新所有指向它的链接"""
        pass
    
    def find_broken_links(self) -> List[LinkInfo]:
        """查找所有断开的链接"""
        pass
    
    def get_link_suggestions(self, prefix: str) -> List[str]:
        """获取链接自动完成建议"""
        pass
    
    def create_link(self, target: str, display_text: Optional[str] = None) -> str:
        """创建 wiki 链接语法"""
        pass
```

### 5. GraphView - 知识图谱视图

```python
from dataclasses import dataclass
from typing import List, Tuple, Optional

@dataclass
class GraphNode:
    """图节点"""
    id: str
    label: str
    x: float
    y: float
    size: float
    color: str
    tags: List[str]

@dataclass
class GraphEdge:
    """图边"""
    source: str
    target: str
    weight: float

@dataclass
class GraphData:
    """图数据"""
    nodes: List[GraphNode]
    edges: List[GraphEdge]

class GraphView:
    """知识图谱视图"""
    
    def __init__(self, link_manager: LinkManager):
        self.link_manager = link_manager
        self._layout_algorithm = "force_directed"
    
    def build_graph(self, documents: List[DocumentInfo]) -> GraphData:
        """构建图数据"""
        pass
    
    def filter_by_tag(self, graph: GraphData, tag: str) -> GraphData:
        """按标签过滤图"""
        pass
    
    def filter_by_depth(self, graph: GraphData, center: str, depth: int) -> GraphData:
        """按深度过滤（以某节点为中心）"""
        pass
    
    def calculate_layout(self, graph: GraphData) -> GraphData:
        """计算节点布局"""
        pass
    
    def get_neighbors(self, node_id: str) -> List[str]:
        """获取相邻节点"""
        pass
    
    def export_to_json(self, graph: GraphData) -> str:
        """导出为 JSON（用于前端渲染）"""
        pass
```

### 6. MetadataParser - 元数据解析器

```python
import yaml
from typing import Optional

class MetadataParser:
    """YAML Frontmatter 解析器"""
    
    FRONTMATTER_PATTERN = re.compile(r'^---\s*\n(.*?)\n---\s*\n', re.DOTALL)
    
    def parse(self, content: str) -> Tuple[dict, str]:
        """解析 frontmatter，返回 (metadata, body)"""
        pass
    
    def extract(self, content: str) -> Optional[dict]:
        """仅提取 frontmatter"""
        pass
    
    def update(self, content: str, metadata: dict) -> str:
        """更新 frontmatter"""
        pass
    
    def create_default(self, title: str) -> dict:
        """创建默认 frontmatter"""
        pass
    
    def validate(self, metadata: dict) -> List[str]:
        """验证 frontmatter，返回错误列表"""
        pass
```

## Data Models

### 数据库 Schema

```sql
-- 文档表
CREATE TABLE documents (
    id INTEGER PRIMARY KEY,
    path TEXT UNIQUE NOT NULL,
    title TEXT,
    content TEXT,
    created_at DATETIME,
    modified_at DATETIME,
    size INTEGER,
    frontmatter TEXT,  -- JSON
    content_hash TEXT
);

-- 标签表
CREATE TABLE tags (
    id INTEGER PRIMARY KEY,
    name TEXT UNIQUE NOT NULL,
    parent_id INTEGER REFERENCES tags(id)
);

-- 文档-标签关联表
CREATE TABLE document_tags (
    document_id INTEGER REFERENCES documents(id),
    tag_id INTEGER REFERENCES tags(id),
    PRIMARY KEY (document_id, tag_id)
);

-- 链接表
CREATE TABLE links (
    id INTEGER PRIMARY KEY,
    source_id INTEGER REFERENCES documents(id),
    target_id INTEGER REFERENCES documents(id),
    target_text TEXT,  -- 原始链接文本
    line_number INTEGER,
    is_broken BOOLEAN DEFAULT FALSE
);

-- 全文搜索索引 (FTS5)
CREATE VIRTUAL TABLE documents_fts USING fts5(
    title, content, tags,
    content='documents',
    content_rowid='id'
);
```

### 配置结构

```python
database_config = {
    'vault': {
        'default_path': '~/Documents/MarkdownVault',
        'watch_interval_ms': 1000,
        'ignore_patterns': ['.*', '_*', 'node_modules']
    },
    'search': {
        'max_results': 100,
        'snippet_length': 150,
        'highlight_tag': '<mark>'
    },
    'graph': {
        'layout': 'force_directed',
        'node_size_by': 'links',  # links, size, none
        'show_orphans': True,
        'max_nodes': 500
    },
    'tags': {
        'hierarchy_separator': '/',
        'auto_create': True
    }
}
```

## Correctness Properties

*A property is a characteristic or behavior that should hold true across all valid executions of a system-essentially, a formal statement about what the system should do. Properties serve as the bridge between human-readable specifications and machine-verifiable correctness guarantees.*

### Property 1: 索引一致性

*For any* 文档库中的文件，索引中的记录应与文件系统状态一致（存在性、修改时间）
**Validates: Requirements 1.3, 1.4**

### Property 2: 搜索结果相关性

*For any* 搜索查询，返回的结果应包含查询关键词，且按相关性降序排列
**Validates: Requirements 2.1, 2.2**

### Property 3: 标签提取正确性

*For any* 包含 #tag 语法的文档，所有标签应被正确提取并索引
**Validates: Requirements 3.1**

### Property 4: 双向链接对称性

*For any* 文档 A 链接到文档 B，则 B 的 backlinks 应包含 A
**Validates: Requirements 4.3, 4.4**

### Property 5: 链接更新传播

*For any* 文档重命名操作，所有指向该文档的链接应被更新
**Validates: Requirements 4.7**

### Property 6: 图数据完整性

*For any* 知识图谱，节点数应等于文档数，边数应等于链接数
**Validates: Requirements 5.1, 5.2**

## Error Handling

```python
class DatabaseError(Exception):
    """数据库相关错误"""
    pass

class VaultError(Exception):
    """文档库相关错误"""
    pass

class IndexError(Exception):
    """索引相关错误"""
    pass
```

## Testing Strategy

### 测试覆盖

| 组件 | 单元测试 | 属性测试 |
| ---- | -------- | -------- |
| VaultManager | ✓ | ✓ (Property 1) |
| SearchEngine | ✓ | ✓ (Property 2) |
| TagManager | ✓ | ✓ (Property 3) |
| LinkManager | ✓ | ✓ (Property 4, 5) |
| GraphView | ✓ | ✓ (Property 6) |

## Dependencies

```
watchdog>=3.0.0  # 文件监控
whoosh>=2.7.4    # 全文搜索（或使用 SQLite FTS5）
pyyaml>=6.0      # YAML 解析
networkx>=3.0    # 图算法
```
