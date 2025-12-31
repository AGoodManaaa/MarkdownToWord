# -*- coding: utf-8 -*-
"""双向链接管理模块"""

from dataclasses import dataclass
from typing import List, Optional, Dict
import re
from pathlib import Path

from .vault import VaultManager, DocumentInfo


@dataclass
class LinkInfo:
    """链接信息"""
    source: str  # 源文档路径
    target: str  # 目标文档路径
    link_text: str  # 链接文本
    line_number: int = 0
    is_broken: bool = False
    
    @property
    def display_text(self) -> str:
        """显示文本"""
        return self.link_text.split('|')[-1] if '|' in self.link_text else self.link_text


class LinkManager:
    """双向链接管理器"""
    
    WIKI_LINK_PATTERN = re.compile(r'\[\[([^\]|]+)(?:\|([^\]]+))?\]\]')
    
    def __init__(self, vault_manager: VaultManager):
        """初始化链接管理器
        
        Args:
            vault_manager: 文档库管理器
        """
        self.vault = vault_manager
        self._link_cache: Dict[str, List[LinkInfo]] = {}
    
    @property
    def _db(self):
        """获取数据库连接"""
        return self.vault._db_conn
    
    def extract_links(self, content: str, source_path: str) -> List[LinkInfo]:
        """从文档内容提取链接
        
        Args:
            content: 文档内容
            source_path: 源文档路径
            
        Returns:
            LinkInfo 列表
        """
        links = []
        lines = content.split('\n')
        
        for line_num, line in enumerate(lines, 1):
            for match in self.WIKI_LINK_PATTERN.finditer(line):
                target = match.group(1)
                display = match.group(2) or target
                
                # 解析目标路径
                resolved = self.resolve_link(target, source_path)
                is_broken = resolved is None
                
                links.append(LinkInfo(
                    source=source_path,
                    target=target,
                    link_text=f"{target}|{display}" if display != target else target,
                    line_number=line_num,
                    is_broken=is_broken
                ))
        
        return links
    
    def get_outgoing_links(self, document_path: str) -> List[LinkInfo]:
        """获取文档的出链
        
        Args:
            document_path: 文档路径
            
        Returns:
            LinkInfo 列表
        """
        cursor = self._db.cursor()
        
        cursor.execute('''
            SELECT l.target_path, l.link_text, l.line_number, l.is_broken
            FROM links l
            JOIN documents d ON d.id = l.source_id
            WHERE d.path = ?
        ''', (document_path,))
        
        links = []
        for row in cursor.fetchall():
            links.append(LinkInfo(
                source=document_path,
                target=row['target_path'],
                link_text=row['link_text'] or row['target_path'],
                line_number=row['line_number'] or 0,
                is_broken=bool(row['is_broken'])
            ))
        
        return links
    
    def get_backlinks(self, document_path: str) -> List[LinkInfo]:
        """获取文档的反向链接
        
        Args:
            document_path: 文档路径
            
        Returns:
            LinkInfo 列表
        """
        cursor = self._db.cursor()
        
        # 获取文档名（不含扩展名）
        doc_name = Path(document_path).stem
        
        # 查找所有指向该文档的链接
        cursor.execute('''
            SELECT d.path as source_path, l.link_text, l.line_number
            FROM links l
            JOIN documents d ON d.id = l.source_id
            WHERE l.target_path = ? OR l.target_path = ? OR l.target_path LIKE ?
        ''', (document_path, doc_name, f'%/{doc_name}'))
        
        links = []
        for row in cursor.fetchall():
            links.append(LinkInfo(
                source=row['source_path'],
                target=document_path,
                link_text=row['link_text'] or doc_name,
                line_number=row['line_number'] or 0,
                is_broken=False
            ))
        
        return links
    
    def resolve_link(self, link_text: str, from_path: str) -> Optional[str]:
        """解析链接目标路径
        
        Args:
            link_text: 链接文本
            from_path: 源文档路径
            
        Returns:
            解析后的目标路径，如果无法解析返回 None
        """
        if not self.vault.current_vault:
            return None
        
        # 移除显示文本部分
        target = link_text.split('|')[0].strip()
        
        # 添加 .md 扩展名（如果没有）
        if not target.endswith('.md'):
            target_with_ext = target + '.md'
        else:
            target_with_ext = target
        
        # 尝试不同的解析策略
        strategies = [
            # 1. 绝对路径（相对于 vault）
            lambda: self.vault.current_vault / target_with_ext,
            # 2. 相对于当前文档的路径
            lambda: (self.vault.current_vault / from_path).parent / target_with_ext,
            # 3. 仅文件名匹配
            lambda: self._find_by_name(target),
        ]
        
        for strategy in strategies:
            try:
                path = strategy()
                if path and path.exists():
                    return str(path.relative_to(self.vault.current_vault))
            except Exception:
                continue
        
        return None
    
    def _find_by_name(self, name: str) -> Optional[Path]:
        """按文件名查找文档
        
        Args:
            name: 文件名（不含扩展名）
            
        Returns:
            文档路径或 None
        """
        if not self.vault.current_vault:
            return None
        
        # 在数据库中查找
        cursor = self._db.cursor()
        cursor.execute('''
            SELECT path FROM documents
            WHERE path LIKE ? OR path LIKE ?
            LIMIT 1
        ''', (f'%/{name}.md', f'{name}.md'))
        
        row = cursor.fetchone()
        if row:
            return self.vault.current_vault / row['path']
        
        return None
    
    def update_links_on_rename(self, old_path: str, new_path: str) -> int:
        """文档重命名时更新所有指向它的链接
        
        Args:
            old_path: 旧路径
            new_path: 新路径
            
        Returns:
            更新的链接数
        """
        cursor = self._db.cursor()
        
        old_name = Path(old_path).stem
        new_name = Path(new_path).stem
        
        # 获取所有指向旧文档的链接
        cursor.execute('''
            SELECT d.id, d.path, d.content, l.id as link_id
            FROM links l
            JOIN documents d ON d.id = l.source_id
            WHERE l.target_path = ? OR l.target_path = ?
        ''', (old_path, old_name))
        
        updated_count = 0
        
        for row in cursor.fetchall():
            content = row['content']
            
            # 替换链接
            new_content = re.sub(
                rf'\[\[{re.escape(old_name)}(\|[^\]]+)?\]\]',
                lambda m: f'[[{new_name}{m.group(1) or ""}]]',
                content
            )
            
            if new_content != content:
                # 更新数据库
                cursor.execute(
                    'UPDATE documents SET content = ? WHERE id = ?',
                    (new_content, row['id'])
                )
                
                # 写回文件
                if self.vault.current_vault:
                    file_path = self.vault.current_vault / row['path']
                    if file_path.exists():
                        with open(file_path, 'w', encoding='utf-8') as f:
                            f.write(new_content)
                
                updated_count += 1
        
        # 更新链接表
        cursor.execute('''
            UPDATE links SET target_path = ?
            WHERE target_path = ? OR target_path = ?
        ''', (new_name, old_path, old_name))
        
        self._db.commit()
        
        return updated_count
    
    def find_broken_links(self) -> List[LinkInfo]:
        """查找所有断开的链接
        
        Returns:
            断开的 LinkInfo 列表
        """
        cursor = self._db.cursor()
        
        cursor.execute('''
            SELECT d.path as source_path, l.target_path, l.link_text, l.line_number
            FROM links l
            JOIN documents d ON d.id = l.source_id
            WHERE l.is_broken = 1
        ''')
        
        links = []
        for row in cursor.fetchall():
            links.append(LinkInfo(
                source=row['source_path'],
                target=row['target_path'],
                link_text=row['link_text'] or row['target_path'],
                line_number=row['line_number'] or 0,
                is_broken=True
            ))
        
        return links
    
    def refresh_broken_links(self) -> None:
        """刷新断开链接状态"""
        cursor = self._db.cursor()
        
        cursor.execute('SELECT id, target_path FROM links')
        
        for row in cursor.fetchall():
            target = row['target_path']
            
            # 检查目标是否存在
            is_broken = True
            
            if self.vault.current_vault:
                # 尝试解析链接
                target_path = target if target.endswith('.md') else f'{target}.md'
                
                # 检查绝对路径
                if (self.vault.current_vault / target_path).exists():
                    is_broken = False
                else:
                    # 检查数据库
                    cursor.execute('''
                        SELECT 1 FROM documents
                        WHERE path = ? OR path LIKE ?
                        LIMIT 1
                    ''', (target_path, f'%/{target_path}'))
                    
                    if cursor.fetchone():
                        is_broken = False
            
            cursor.execute(
                'UPDATE links SET is_broken = ? WHERE id = ?',
                (is_broken, row['id'])
            )
        
        self._db.commit()
    
    def get_link_suggestions(self, prefix: str, limit: int = 10) -> List[str]:
        """获取链接自动完成建议
        
        Args:
            prefix: 输入前缀
            limit: 建议数量限制
            
        Returns:
            文档名建议列表
        """
        cursor = self._db.cursor()
        
        # 搜索匹配的文档
        pattern = f'%{prefix}%'
        cursor.execute('''
            SELECT path, title FROM documents
            WHERE path LIKE ? OR title LIKE ?
            ORDER BY modified_at DESC
            LIMIT ?
        ''', (pattern, pattern, limit))
        
        suggestions = []
        for row in cursor.fetchall():
            # 返回文件名（不含扩展名）
            name = Path(row['path']).stem
            suggestions.append(name)
        
        return suggestions
    
    def create_link(self, target: str, display_text: Optional[str] = None) -> str:
        """创建 wiki 链接语法
        
        Args:
            target: 目标文档
            display_text: 显示文本（可选）
            
        Returns:
            wiki 链接字符串
        """
        if display_text and display_text != target:
            return f'[[{target}|{display_text}]]'
        return f'[[{target}]]'
    
    def get_link_graph_data(self) -> Dict:
        """获取链接图数据
        
        Returns:
            包含节点和边的字典
        """
        cursor = self._db.cursor()
        
        # 获取所有文档作为节点
        cursor.execute('SELECT path, title FROM documents')
        nodes = {row['path']: row['title'] for row in cursor.fetchall()}
        
        # 获取所有链接作为边
        cursor.execute('''
            SELECT d.path as source, l.target_path as target
            FROM links l
            JOIN documents d ON d.id = l.source_id
            WHERE l.is_broken = 0
        ''')
        
        edges = []
        for row in cursor.fetchall():
            source = row['source']
            target = row['target']
            
            # 解析目标路径
            if not target.endswith('.md'):
                target = f'{target}.md'
            
            if target in nodes or any(target in p for p in nodes):
                edges.append((source, target))
        
        return {
            'nodes': nodes,
            'edges': edges
        }
