# -*- coding: utf-8 -*-
"""文档库管理模块"""

from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Callable, Dict, Any
from datetime import datetime
import os
import sqlite3
import json
import hashlib
import threading

try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler, FileModifiedEvent, FileCreatedEvent, FileDeletedEvent
    WATCHDOG_AVAILABLE = True
except ImportError:
    WATCHDOG_AVAILABLE = False
    Observer = None
    FileSystemEventHandler = object


class VaultError(Exception):
    """文档库相关错误"""
    pass


@dataclass
class VaultInfo:
    """文档库信息"""
    path: Path
    name: str
    file_count: int = 0
    last_indexed: Optional[datetime] = None
    size_bytes: int = 0


@dataclass
class DocumentInfo:
    """文档信息"""
    path: Path
    title: str
    created: Optional[datetime] = None
    modified: Optional[datetime] = None
    size: int = 0
    tags: List[str] = field(default_factory=list)
    links: List[str] = field(default_factory=list)
    backlinks: List[str] = field(default_factory=list)
    frontmatter: Dict[str, Any] = field(default_factory=dict)
    content_hash: str = ""
    
    @property
    def filename(self) -> str:
        """文件名（不含扩展名）"""
        return self.path.stem
    
    @property
    def relative_path(self) -> str:
        """相对路径字符串"""
        return str(self.path)


class VaultFileHandler(FileSystemEventHandler):
    """文件变化处理器"""
    
    def __init__(self, callback: Callable[[str, str], None]):
        self.callback = callback
        self._debounce_timers: Dict[str, threading.Timer] = {}
        self._lock = threading.Lock()
    
    def _debounced_callback(self, path: str, event_type: str):
        """防抖回调"""
        with self._lock:
            if path in self._debounce_timers:
                self._debounce_timers[path].cancel()
            
            timer = threading.Timer(0.5, lambda: self.callback(path, event_type))
            self._debounce_timers[path] = timer
            timer.start()
    
    def on_modified(self, event):
        if not event.is_directory and event.src_path.endswith('.md'):
            self._debounced_callback(event.src_path, 'modified')
    
    def on_created(self, event):
        if not event.is_directory and event.src_path.endswith('.md'):
            self._debounced_callback(event.src_path, 'created')
    
    def on_deleted(self, event):
        if not event.is_directory and event.src_path.endswith('.md'):
            self._debounced_callback(event.src_path, 'deleted')
    
    def on_moved(self, event):
        if not event.is_directory and event.src_path.endswith('.md'):
            self._debounced_callback(event.src_path, 'deleted')
            if event.dest_path.endswith('.md'):
                self._debounced_callback(event.dest_path, 'created')


class VaultManager:
    """文档库管理器"""
    
    IGNORE_PATTERNS = {'.git', '.obsidian', '.trash', 'node_modules', '__pycache__'}
    
    def __init__(self, db_path: Optional[str] = None):
        """初始化文档库管理器
        
        Args:
            db_path: 数据库文件路径，如果为 None 则使用内存数据库
        """
        self.db_path = db_path
        self.current_vault: Optional[Path] = None
        self._db_conn: Optional[sqlite3.Connection] = None
        self._file_watcher = None
        self._observer = None
        self._on_file_change: Optional[Callable[[str, str], None]] = None
        
        self._init_database()
    
    def _init_database(self) -> None:
        """初始化数据库"""
        if self.db_path:
            self._db_conn = sqlite3.connect(self.db_path, check_same_thread=False)
        else:
            self._db_conn = sqlite3.connect(':memory:', check_same_thread=False)
        
        self._db_conn.row_factory = sqlite3.Row
        self._create_tables()
    
    def _create_tables(self) -> None:
        """创建数据库表"""
        cursor = self._db_conn.cursor()
        
        # 文档表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY,
                path TEXT UNIQUE NOT NULL,
                title TEXT,
                content TEXT,
                created_at DATETIME,
                modified_at DATETIME,
                size INTEGER,
                frontmatter TEXT,
                content_hash TEXT
            )
        ''')
        
        # 标签表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS tags (
                id INTEGER PRIMARY KEY,
                name TEXT UNIQUE NOT NULL,
                parent_id INTEGER REFERENCES tags(id)
            )
        ''')
        
        # 文档-标签关联表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS document_tags (
                document_id INTEGER REFERENCES documents(id),
                tag_id INTEGER REFERENCES tags(id),
                PRIMARY KEY (document_id, tag_id)
            )
        ''')
        
        # 链接表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS links (
                id INTEGER PRIMARY KEY,
                source_id INTEGER REFERENCES documents(id),
                target_path TEXT,
                link_text TEXT,
                line_number INTEGER,
                is_broken BOOLEAN DEFAULT FALSE
            )
        ''')
        
        # 全文搜索索引 (FTS5)
        cursor.execute('''
            CREATE VIRTUAL TABLE IF NOT EXISTS documents_fts USING fts5(
                title, content, tags,
                content='documents',
                content_rowid='id'
            )
        ''')
        
        self._db_conn.commit()
    
    def open_vault(self, path: str) -> VaultInfo:
        """打开文档库
        
        Args:
            path: 文档库路径
            
        Returns:
            VaultInfo 文档库信息
            
        Raises:
            VaultError: 路径不存在或不可访问
        """
        vault_path = Path(path).resolve()
        
        if not vault_path.exists():
            raise VaultError(f"路径不存在: {path}")
        
        if not vault_path.is_dir():
            raise VaultError(f"路径不是目录: {path}")
        
        self.current_vault = vault_path
        
        # 扫描并索引文档
        documents = self.scan_vault()
        
        # 计算总大小
        total_size = sum(d.size for d in documents)
        
        return VaultInfo(
            path=vault_path,
            name=vault_path.name,
            file_count=len(documents),
            last_indexed=datetime.now(),
            size_bytes=total_size
        )
    
    def create_vault(self, path: str, name: str) -> VaultInfo:
        """创建新文档库
        
        Args:
            path: 文档库路径
            name: 文档库名称
            
        Returns:
            VaultInfo 文档库信息
        """
        vault_path = Path(path).resolve()
        
        if not vault_path.exists():
            vault_path.mkdir(parents=True)
        
        self.current_vault = vault_path
        
        return VaultInfo(
            path=vault_path,
            name=name,
            file_count=0,
            last_indexed=datetime.now(),
            size_bytes=0
        )
    
    def scan_vault(self) -> List[DocumentInfo]:
        """扫描文档库中的所有文件
        
        Returns:
            DocumentInfo 列表
        """
        if self.current_vault is None:
            return []
        
        documents = []
        
        for md_file in self.current_vault.rglob('*.md'):
            # 跳过忽略的目录
            if any(part in self.IGNORE_PATTERNS for part in md_file.parts):
                continue
            
            try:
                doc = self._index_document(md_file)
                if doc:
                    documents.append(doc)
            except Exception:
                continue
        
        return documents
    
    def _index_document(self, file_path: Path) -> Optional[DocumentInfo]:
        """索引单个文档
        
        Args:
            file_path: 文档路径
            
        Returns:
            DocumentInfo 或 None
        """
        try:
            stat = file_path.stat()
            
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 计算内容哈希
            content_hash = hashlib.md5(content.encode()).hexdigest()
            
            # 解析 frontmatter
            frontmatter, body = self._parse_frontmatter(content)
            
            # 提取标题
            title = frontmatter.get('title', '')
            if not title:
                # 从内容中提取第一个标题
                for line in body.split('\n'):
                    if line.startswith('# '):
                        title = line[2:].strip()
                        break
            if not title:
                title = file_path.stem
            
            # 提取标签
            tags = self._extract_tags(content)
            if 'tags' in frontmatter:
                fm_tags = frontmatter['tags']
                if isinstance(fm_tags, list):
                    tags.extend(fm_tags)
                elif isinstance(fm_tags, str):
                    tags.append(fm_tags)
            
            # 提取链接
            links = self._extract_links(content)
            
            # 相对路径
            rel_path = file_path.relative_to(self.current_vault)
            
            doc = DocumentInfo(
                path=rel_path,
                title=title,
                created=datetime.fromtimestamp(stat.st_ctime),
                modified=datetime.fromtimestamp(stat.st_mtime),
                size=stat.st_size,
                tags=list(set(tags)),
                links=links,
                frontmatter=frontmatter,
                content_hash=content_hash
            )
            
            # 保存到数据库
            self._save_document_to_db(doc, content)
            
            return doc
            
        except Exception:
            return None
    
    def _parse_frontmatter(self, content: str) -> tuple:
        """解析 YAML frontmatter
        
        Args:
            content: 文档内容
            
        Returns:
            (frontmatter_dict, body)
        """
        import re
        
        pattern = re.compile(r'^---\s*\n(.*?)\n---\s*\n', re.DOTALL)
        match = pattern.match(content)
        
        if not match:
            return {}, content
        
        try:
            import yaml
            frontmatter = yaml.safe_load(match.group(1)) or {}
            body = content[match.end():]
            return frontmatter, body
        except Exception:
            return {}, content
    
    def _extract_tags(self, content: str) -> List[str]:
        """从内容中提取标签
        
        Args:
            content: 文档内容
            
        Returns:
            标签列表
        """
        import re
        
        # 匹配 #tag 格式（不在代码块中）
        pattern = re.compile(r'(?<!\S)#([a-zA-Z\u4e00-\u9fff][a-zA-Z0-9\u4e00-\u9fff_/-]*)')
        
        tags = []
        for match in pattern.finditer(content):
            tag = match.group(1)
            # 排除常见的非标签模式
            if not tag.startswith(('#', '!')):
                tags.append(tag)
        
        return tags
    
    def _extract_links(self, content: str) -> List[str]:
        """从内容中提取 wiki 链接
        
        Args:
            content: 文档内容
            
        Returns:
            链接目标列表
        """
        import re
        
        # 匹配 [[link]] 或 [[link|display]] 格式
        pattern = re.compile(r'\[\[([^\]|]+)(?:\|[^\]]+)?\]\]')
        
        links = []
        for match in pattern.finditer(content):
            links.append(match.group(1))
        
        return links
    
    def _save_document_to_db(self, doc: DocumentInfo, content: str) -> None:
        """保存文档到数据库
        
        Args:
            doc: 文档信息
            content: 文档内容
        """
        cursor = self._db_conn.cursor()
        
        # 插入或更新文档
        cursor.execute('''
            INSERT OR REPLACE INTO documents 
            (path, title, content, created_at, modified_at, size, frontmatter, content_hash)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            str(doc.path),
            doc.title,
            content,
            doc.created.isoformat() if doc.created else None,
            doc.modified.isoformat() if doc.modified else None,
            doc.size,
            json.dumps(doc.frontmatter, ensure_ascii=False),
            doc.content_hash
        ))
        
        doc_id = cursor.lastrowid
        
        # 更新标签
        cursor.execute('DELETE FROM document_tags WHERE document_id = ?', (doc_id,))
        
        for tag in doc.tags:
            # 确保标签存在
            cursor.execute('INSERT OR IGNORE INTO tags (name) VALUES (?)', (tag,))
            cursor.execute('SELECT id FROM tags WHERE name = ?', (tag,))
            tag_id = cursor.fetchone()[0]
            
            cursor.execute(
                'INSERT OR IGNORE INTO document_tags (document_id, tag_id) VALUES (?, ?)',
                (doc_id, tag_id)
            )
        
        # 更新链接
        cursor.execute('DELETE FROM links WHERE source_id = ?', (doc_id,))
        
        for link in doc.links:
            cursor.execute('''
                INSERT INTO links (source_id, target_path, link_text)
                VALUES (?, ?, ?)
            ''', (doc_id, link, link))
        
        # 更新全文索引
        cursor.execute('''
            INSERT OR REPLACE INTO documents_fts (rowid, title, content, tags)
            VALUES (?, ?, ?, ?)
        ''', (doc_id, doc.title, content, ' '.join(doc.tags)))
        
        self._db_conn.commit()
    
    def watch_changes(self, callback: Callable[[str, str], None]) -> None:
        """监控文件变化
        
        Args:
            callback: 变化回调函数 (path, event_type)
        """
        if not WATCHDOG_AVAILABLE:
            return
        
        if self.current_vault is None:
            return
        
        self._on_file_change = callback
        
        self._file_watcher = VaultFileHandler(self._handle_file_change)
        self._observer = Observer()
        self._observer.schedule(
            self._file_watcher,
            str(self.current_vault),
            recursive=True
        )
        self._observer.start()
    
    def _handle_file_change(self, path: str, event_type: str) -> None:
        """处理文件变化"""
        if event_type == 'deleted':
            self._remove_document_from_db(path)
        else:
            file_path = Path(path)
            if file_path.exists():
                self._index_document(file_path)
        
        if self._on_file_change:
            self._on_file_change(path, event_type)
    
    def _remove_document_from_db(self, path: str) -> None:
        """从数据库移除文档"""
        cursor = self._db_conn.cursor()
        
        # 获取文档 ID
        cursor.execute('SELECT id FROM documents WHERE path = ?', (path,))
        row = cursor.fetchone()
        
        if row:
            doc_id = row[0]
            cursor.execute('DELETE FROM document_tags WHERE document_id = ?', (doc_id,))
            cursor.execute('DELETE FROM links WHERE source_id = ?', (doc_id,))
            cursor.execute('DELETE FROM documents WHERE id = ?', (doc_id,))
            cursor.execute('DELETE FROM documents_fts WHERE rowid = ?', (doc_id,))
            self._db_conn.commit()
    
    def stop_watching(self) -> None:
        """停止监控"""
        if self._observer:
            self._observer.stop()
            self._observer.join()
            self._observer = None
        self._file_watcher = None
    
    def get_document(self, path: str) -> Optional[DocumentInfo]:
        """获取文档信息
        
        Args:
            path: 文档路径
            
        Returns:
            DocumentInfo 或 None
        """
        cursor = self._db_conn.cursor()
        cursor.execute('''
            SELECT path, title, created_at, modified_at, size, frontmatter, content_hash
            FROM documents WHERE path = ?
        ''', (path,))
        
        row = cursor.fetchone()
        if not row:
            return None
        
        # 获取标签
        cursor.execute('''
            SELECT t.name FROM tags t
            JOIN document_tags dt ON t.id = dt.tag_id
            JOIN documents d ON d.id = dt.document_id
            WHERE d.path = ?
        ''', (path,))
        tags = [r[0] for r in cursor.fetchall()]
        
        # 获取链接
        cursor.execute('''
            SELECT target_path FROM links l
            JOIN documents d ON d.id = l.source_id
            WHERE d.path = ?
        ''', (path,))
        links = [r[0] for r in cursor.fetchall()]
        
        # 获取反向链接
        cursor.execute('''
            SELECT d.path FROM documents d
            JOIN links l ON d.id = l.source_id
            WHERE l.target_path = ?
        ''', (path,))
        backlinks = [r[0] for r in cursor.fetchall()]
        
        return DocumentInfo(
            path=Path(row['path']),
            title=row['title'],
            created=datetime.fromisoformat(row['created_at']) if row['created_at'] else None,
            modified=datetime.fromisoformat(row['modified_at']) if row['modified_at'] else None,
            size=row['size'],
            tags=tags,
            links=links,
            backlinks=backlinks,
            frontmatter=json.loads(row['frontmatter']) if row['frontmatter'] else {},
            content_hash=row['content_hash']
        )
    
    def get_all_documents(self) -> List[DocumentInfo]:
        """获取所有文档"""
        cursor = self._db_conn.cursor()
        cursor.execute('SELECT path FROM documents')
        
        documents = []
        for row in cursor.fetchall():
            doc = self.get_document(row['path'])
            if doc:
                documents.append(doc)
        
        return documents
    
    def get_recent_documents(self, limit: int = 10) -> List[DocumentInfo]:
        """获取最近修改的文档
        
        Args:
            limit: 返回数量限制
            
        Returns:
            DocumentInfo 列表
        """
        cursor = self._db_conn.cursor()
        cursor.execute('''
            SELECT path FROM documents
            ORDER BY modified_at DESC
            LIMIT ?
        ''', (limit,))
        
        documents = []
        for row in cursor.fetchall():
            doc = self.get_document(row['path'])
            if doc:
                documents.append(doc)
        
        return documents
    
    def close(self) -> None:
        """关闭文档库"""
        self.stop_watching()
        if self._db_conn:
            self._db_conn.close()
            self._db_conn = None
