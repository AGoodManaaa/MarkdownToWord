# -*- coding: utf-8 -*-
"""标签管理模块"""

from dataclasses import dataclass, field
from typing import List, Optional, Dict
import re

from .vault import VaultManager, DocumentInfo


@dataclass
class TagInfo:
    """标签信息"""
    name: str
    count: int = 0
    parent: Optional[str] = None
    children: List[str] = field(default_factory=list)
    
    @property
    def is_hierarchical(self) -> bool:
        """是否为层级标签"""
        return '/' in self.name
    
    @property
    def parts(self) -> List[str]:
        """标签层级部分"""
        return self.name.split('/')


class TagManager:
    """标签管理器"""
    
    TAG_PATTERN = re.compile(r'(?<!\S)#([a-zA-Z\u4e00-\u9fff][a-zA-Z0-9\u4e00-\u9fff_/-]*)')
    HIERARCHY_SEPARATOR = '/'
    
    def __init__(self, vault_manager: VaultManager):
        """初始化标签管理器
        
        Args:
            vault_manager: 文档库管理器
        """
        self.vault = vault_manager
    
    @property
    def _db(self):
        """获取数据库连接"""
        return self.vault._db_conn
    
    def extract_tags(self, content: str) -> List[str]:
        """从文档内容提取标签
        
        Args:
            content: 文档内容
            
        Returns:
            标签列表
        """
        tags = []
        
        for match in self.TAG_PATTERN.finditer(content):
            tag = match.group(1)
            # 排除常见的非标签模式（如标题标记）
            if not tag.startswith('#'):
                tags.append(tag)
        
        return list(set(tags))
    
    def get_all_tags(self) -> List[TagInfo]:
        """获取所有标签
        
        Returns:
            TagInfo 列表
        """
        cursor = self._db.cursor()
        
        # 获取所有标签及其文档计数
        cursor.execute('''
            SELECT t.name, COUNT(dt.document_id) as count
            FROM tags t
            LEFT JOIN document_tags dt ON t.id = dt.tag_id
            GROUP BY t.id
            ORDER BY count DESC, t.name
        ''')
        
        tags_dict: Dict[str, TagInfo] = {}
        
        for row in cursor.fetchall():
            name = row['name']
            count = row['count']
            
            # 解析层级
            parts = name.split(self.HIERARCHY_SEPARATOR)
            parent = None
            
            if len(parts) > 1:
                parent = self.HIERARCHY_SEPARATOR.join(parts[:-1])
            
            tags_dict[name] = TagInfo(
                name=name,
                count=count,
                parent=parent
            )
        
        # 建立父子关系
        for tag in tags_dict.values():
            if tag.parent and tag.parent in tags_dict:
                tags_dict[tag.parent].children.append(tag.name)
        
        return list(tags_dict.values())
    
    def get_documents_by_tag(self, tag: str) -> List[DocumentInfo]:
        """获取包含指定标签的文档
        
        Args:
            tag: 标签名
            
        Returns:
            DocumentInfo 列表
        """
        cursor = self._db.cursor()
        
        # 支持层级标签匹配（匹配标签及其子标签）
        cursor.execute('''
            SELECT DISTINCT d.path FROM documents d
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
    
    def rename_tag(self, old_name: str, new_name: str) -> int:
        """重命名标签
        
        Args:
            old_name: 旧标签名
            new_name: 新标签名
            
        Returns:
            受影响的文档数
        """
        cursor = self._db.cursor()
        
        # 获取受影响的文档
        cursor.execute('''
            SELECT d.id, d.path, d.content FROM documents d
            JOIN document_tags dt ON d.id = dt.document_id
            JOIN tags t ON t.id = dt.tag_id
            WHERE t.name = ? OR t.name LIKE ?
        ''', (old_name, f'{old_name}/%'))
        
        affected_docs = cursor.fetchall()
        
        # 更新文档内容中的标签
        for doc in affected_docs:
            content = doc['content']
            
            # 替换标签
            new_content = re.sub(
                rf'(?<!\S)#{re.escape(old_name)}(?=\s|$)',
                f'#{new_name}',
                content
            )
            
            # 替换子标签
            new_content = re.sub(
                rf'(?<!\S)#{re.escape(old_name)}/',
                f'#{new_name}/',
                new_content
            )
            
            if new_content != content:
                # 更新数据库
                cursor.execute(
                    'UPDATE documents SET content = ? WHERE id = ?',
                    (new_content, doc['id'])
                )
                
                # 写回文件
                if self.vault.current_vault:
                    file_path = self.vault.current_vault / doc['path']
                    if file_path.exists():
                        with open(file_path, 'w', encoding='utf-8') as f:
                            f.write(new_content)
        
        # 更新标签表
        cursor.execute(
            'UPDATE tags SET name = ? WHERE name = ?',
            (new_name, old_name)
        )
        
        # 更新子标签
        cursor.execute('''
            UPDATE tags SET name = REPLACE(name, ?, ?)
            WHERE name LIKE ?
        ''', (f'{old_name}/', f'{new_name}/', f'{old_name}/%'))
        
        self._db.commit()
        
        return len(affected_docs)
    
    def delete_tag(self, tag: str) -> int:
        """删除标签（从文档中移除）
        
        Args:
            tag: 标签名
            
        Returns:
            受影响的文档数
        """
        cursor = self._db.cursor()
        
        # 获取受影响的文档
        cursor.execute('''
            SELECT d.id, d.path, d.content FROM documents d
            JOIN document_tags dt ON d.id = dt.document_id
            JOIN tags t ON t.id = dt.tag_id
            WHERE t.name = ?
        ''', (tag,))
        
        affected_docs = cursor.fetchall()
        
        # 从文档内容中移除标签
        for doc in affected_docs:
            content = doc['content']
            
            # 移除标签
            new_content = re.sub(
                rf'(?<!\S)#{re.escape(tag)}(?=\s|$)',
                '',
                content
            )
            
            if new_content != content:
                cursor.execute(
                    'UPDATE documents SET content = ? WHERE id = ?',
                    (new_content, doc['id'])
                )
                
                # 写回文件
                if self.vault.current_vault:
                    file_path = self.vault.current_vault / doc['path']
                    if file_path.exists():
                        with open(file_path, 'w', encoding='utf-8') as f:
                            f.write(new_content)
        
        # 从数据库删除标签
        cursor.execute('SELECT id FROM tags WHERE name = ?', (tag,))
        tag_row = cursor.fetchone()
        
        if tag_row:
            cursor.execute('DELETE FROM document_tags WHERE tag_id = ?', (tag_row['id'],))
            cursor.execute('DELETE FROM tags WHERE id = ?', (tag_row['id'],))
        
        self._db.commit()
        
        return len(affected_docs)
    
    def merge_tags(self, source: str, target: str) -> int:
        """合并标签
        
        Args:
            source: 源标签（将被删除）
            target: 目标标签
            
        Returns:
            受影响的文档数
        """
        # 先重命名，再删除重复
        affected = self.rename_tag(source, target)
        
        # 清理重复的标签关联
        cursor = self._db.cursor()
        cursor.execute('''
            DELETE FROM document_tags
            WHERE rowid NOT IN (
                SELECT MIN(rowid) FROM document_tags
                GROUP BY document_id, tag_id
            )
        ''')
        
        self._db.commit()
        
        return affected
    
    def get_tag_hierarchy(self) -> Dict[str, List[str]]:
        """获取标签层级结构
        
        Returns:
            层级结构字典 {parent: [children]}
        """
        tags = self.get_all_tags()
        
        hierarchy: Dict[str, List[str]] = {'': []}  # 根级别
        
        for tag in tags:
            if tag.parent:
                if tag.parent not in hierarchy:
                    hierarchy[tag.parent] = []
                hierarchy[tag.parent].append(tag.name)
            else:
                hierarchy[''].append(tag.name)
        
        return hierarchy
    
    def get_tag_cloud(self, limit: int = 50) -> List[TagInfo]:
        """获取标签云数据
        
        Args:
            limit: 返回数量限制
            
        Returns:
            按使用频率排序的 TagInfo 列表
        """
        cursor = self._db.cursor()
        
        cursor.execute('''
            SELECT t.name, COUNT(dt.document_id) as count
            FROM tags t
            JOIN document_tags dt ON t.id = dt.tag_id
            GROUP BY t.id
            HAVING count > 0
            ORDER BY count DESC
            LIMIT ?
        ''', (limit,))
        
        return [
            TagInfo(name=row['name'], count=row['count'])
            for row in cursor.fetchall()
        ]
