# -*- coding: utf-8 -*-
"""全文搜索引擎模块"""

from dataclasses import dataclass, field
from typing import List, Optional, Tuple
import re
import sqlite3

from .vault import VaultManager, DocumentInfo


@dataclass
class SearchMatch:
    """匹配详情"""
    line_number: int
    content: str
    highlight_ranges: List[Tuple[int, int]] = field(default_factory=list)  # (start, end)


@dataclass
class SearchResult:
    """搜索结果"""
    document: DocumentInfo
    score: float
    matches: List[SearchMatch] = field(default_factory=list)
    
    @property
    def snippet(self) -> str:
        """获取匹配片段"""
        if self.matches:
            return self.matches[0].content[:150]
        return ""


class SearchEngine:
    """全文搜索引擎"""
    
    def __init__(self, vault_manager: VaultManager):
        """初始化搜索引擎
        
        Args:
            vault_manager: 文档库管理器
        """
        self.vault = vault_manager
        self._snippet_length = 150
    
    @property
    def _db(self) -> sqlite3.Connection:
        """获取数据库连接"""
        return self.vault._db_conn
    
    def search(self, query: str, limit: int = 50) -> List[SearchResult]:
        """执行搜索
        
        Args:
            query: 搜索查询
            limit: 结果数量限制
            
        Returns:
            SearchResult 列表
        """
        if not query.strip():
            return []
        
        # 解析查询
        parsed = self.parse_query(query)
        
        # 构建 FTS5 查询
        fts_query = self._build_fts_query(parsed)
        
        cursor = self._db.cursor()
        
        try:
            # 执行搜索
            cursor.execute('''
                SELECT d.path, d.content, bm25(documents_fts) as score
                FROM documents_fts
                JOIN documents d ON documents_fts.rowid = d.id
                WHERE documents_fts MATCH ?
                ORDER BY score
                LIMIT ?
            ''', (fts_query, limit))
            
            results = []
            for row in cursor.fetchall():
                doc = self.vault.get_document(row['path'])
                if doc:
                    # 提取匹配片段
                    matches = self._extract_matches(row['content'], query)
                    
                    results.append(SearchResult(
                        document=doc,
                        score=abs(row['score']),
                        matches=matches
                    ))
            
            return results
            
        except sqlite3.OperationalError:
            # FTS 查询失败，回退到简单搜索
            return self._simple_search(query, limit)
    
    def _simple_search(self, query: str, limit: int) -> List[SearchResult]:
        """简单搜索（回退方案）"""
        cursor = self._db.cursor()
        
        # 使用 LIKE 进行简单搜索
        pattern = f'%{query}%'
        cursor.execute('''
            SELECT path, content FROM documents
            WHERE content LIKE ? OR title LIKE ?
            LIMIT ?
        ''', (pattern, pattern, limit))
        
        results = []
        for row in cursor.fetchall():
            doc = self.vault.get_document(row['path'])
            if doc:
                matches = self._extract_matches(row['content'], query)
                results.append(SearchResult(
                    document=doc,
                    score=1.0,
                    matches=matches
                ))
        
        return results
    
    def search_by_tag(self, tag: str) -> List[DocumentInfo]:
        """按标签搜索
        
        Args:
            tag: 标签名
            
        Returns:
            DocumentInfo 列表
        """
        cursor = self._db.cursor()
        
        cursor.execute('''
            SELECT d.path FROM documents d
            JOIN document_tags dt ON d.id = dt.document_id
            JOIN tags t ON t.id = dt.tag_id
            WHERE t.name = ? OR t.name LIKE ?
        ''', (tag, f'{tag}/%'))
        
        documents = []
        for row in cursor.fetchall():
            doc = self.vault.get_document(row['path'])
            if doc:
                documents.append(doc)
        
        return documents
    
    def search_by_filename(self, pattern: str) -> List[DocumentInfo]:
        """按文件名搜索
        
        Args:
            pattern: 文件名模式
            
        Returns:
            DocumentInfo 列表
        """
        cursor = self._db.cursor()
        
        # 转换通配符
        sql_pattern = pattern.replace('*', '%').replace('?', '_')
        if not sql_pattern.startswith('%'):
            sql_pattern = '%' + sql_pattern
        if not sql_pattern.endswith('%'):
            sql_pattern = sql_pattern + '%'
        
        cursor.execute('''
            SELECT path FROM documents
            WHERE path LIKE ?
        ''', (sql_pattern,))
        
        documents = []
        for row in cursor.fetchall():
            doc = self.vault.get_document(row['path'])
            if doc:
                documents.append(doc)
        
        return documents
    
    def parse_query(self, query: str) -> dict:
        """解析高级查询语法
        
        支持:
        - AND: 空格分隔的词默认为 AND
        - OR: 使用 OR 关键字
        - NOT: 使用 - 前缀
        - 引号: 精确匹配短语
        
        Args:
            query: 查询字符串
            
        Returns:
            解析后的查询结构
        """
        result = {
            'must': [],      # AND 条件
            'should': [],    # OR 条件
            'must_not': [],  # NOT 条件
            'phrases': []    # 精确短语
        }
        
        # 提取引号中的短语
        phrase_pattern = re.compile(r'"([^"]+)"')
        for match in phrase_pattern.finditer(query):
            result['phrases'].append(match.group(1))
        
        # 移除已处理的短语
        query = phrase_pattern.sub('', query)
        
        # 分割词
        tokens = query.split()
        
        i = 0
        while i < len(tokens):
            token = tokens[i]
            
            if token.upper() == 'OR' and i + 1 < len(tokens):
                # OR 操作
                if result['must']:
                    result['should'].append(result['must'].pop())
                result['should'].append(tokens[i + 1])
                i += 2
            elif token.startswith('-'):
                # NOT 操作
                result['must_not'].append(token[1:])
                i += 1
            elif token.upper() not in ('AND', 'OR', 'NOT'):
                # 普通词（AND）
                result['must'].append(token)
                i += 1
            else:
                i += 1
        
        return result
    
    def _build_fts_query(self, parsed: dict) -> str:
        """构建 FTS5 查询字符串
        
        Args:
            parsed: 解析后的查询结构
            
        Returns:
            FTS5 查询字符串
        """
        parts = []
        
        # AND 条件
        for term in parsed['must']:
            parts.append(f'"{term}"')
        
        # OR 条件
        if parsed['should']:
            or_part = ' OR '.join(f'"{t}"' for t in parsed['should'])
            parts.append(f'({or_part})')
        
        # NOT 条件
        for term in parsed['must_not']:
            parts.append(f'NOT "{term}"')
        
        # 精确短语
        for phrase in parsed['phrases']:
            parts.append(f'"{phrase}"')
        
        return ' '.join(parts) if parts else '*'
    
    def _extract_matches(self, content: str, query: str) -> List[SearchMatch]:
        """提取匹配片段
        
        Args:
            content: 文档内容
            query: 搜索查询
            
        Returns:
            SearchMatch 列表
        """
        matches = []
        
        # 获取搜索词
        terms = re.findall(r'\w+', query.lower())
        if not terms:
            return matches
        
        lines = content.split('\n')
        
        for line_num, line in enumerate(lines, 1):
            line_lower = line.lower()
            
            # 检查是否包含任何搜索词
            found_ranges = []
            for term in terms:
                start = 0
                while True:
                    pos = line_lower.find(term, start)
                    if pos == -1:
                        break
                    found_ranges.append((pos, pos + len(term)))
                    start = pos + 1
            
            if found_ranges:
                # 合并重叠的范围
                found_ranges.sort()
                merged = []
                for start, end in found_ranges:
                    if merged and start <= merged[-1][1]:
                        merged[-1] = (merged[-1][0], max(merged[-1][1], end))
                    else:
                        merged.append((start, end))
                
                matches.append(SearchMatch(
                    line_number=line_num,
                    content=line[:self._snippet_length],
                    highlight_ranges=merged
                ))
                
                if len(matches) >= 5:  # 最多返回 5 个匹配
                    break
        
        return matches
    
    def get_suggestions(self, prefix: str, limit: int = 10) -> List[str]:
        """获取搜索建议
        
        Args:
            prefix: 输入前缀
            limit: 建议数量限制
            
        Returns:
            建议列表
        """
        if not prefix or len(prefix) < 2:
            return []
        
        cursor = self._db.cursor()
        
        # 从标题中获取建议
        cursor.execute('''
            SELECT DISTINCT title FROM documents
            WHERE title LIKE ?
            LIMIT ?
        ''', (f'{prefix}%', limit))
        
        suggestions = [row['title'] for row in cursor.fetchall()]
        
        # 从标签中获取建议
        cursor.execute('''
            SELECT DISTINCT name FROM tags
            WHERE name LIKE ?
            LIMIT ?
        ''', (f'{prefix}%', limit - len(suggestions)))
        
        suggestions.extend(f'#{row["name"]}' for row in cursor.fetchall())
        
        return suggestions[:limit]
    
    def rebuild_index(self) -> None:
        """重建搜索索引"""
        cursor = self._db.cursor()
        
        # 清空 FTS 表
        cursor.execute('DELETE FROM documents_fts')
        
        # 重新索引所有文档
        cursor.execute('SELECT id, title, content FROM documents')
        
        for row in cursor.fetchall():
            # 获取标签
            cursor.execute('''
                SELECT t.name FROM tags t
                JOIN document_tags dt ON t.id = dt.tag_id
                WHERE dt.document_id = ?
            ''', (row['id'],))
            tags = ' '.join(r['name'] for r in cursor.fetchall())
            
            cursor.execute('''
                INSERT INTO documents_fts (rowid, title, content, tags)
                VALUES (?, ?, ?, ?)
            ''', (row['id'], row['title'], row['content'], tags))
        
        self._db.commit()
