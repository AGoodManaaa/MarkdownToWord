# -*- coding: utf-8 -*-
"""元数据解析模块"""

import re
from typing import Tuple, Optional, List, Dict, Any
from datetime import datetime


class MetadataParser:
    """YAML Frontmatter 解析器"""
    
    FRONTMATTER_PATTERN = re.compile(r'^---\s*\n(.*?)\n---\s*\n', re.DOTALL)
    
    def parse(self, content: str) -> Tuple[Dict[str, Any], str]:
        """解析 frontmatter
        
        Args:
            content: 文档内容
            
        Returns:
            (metadata_dict, body) 元组
        """
        match = self.FRONTMATTER_PATTERN.match(content)
        
        if not match:
            return {}, content
        
        try:
            import yaml
            frontmatter = yaml.safe_load(match.group(1))
            if frontmatter is None:
                frontmatter = {}
            body = content[match.end():]
            return frontmatter, body
        except Exception:
            return {}, content
    
    def extract(self, content: str) -> Optional[Dict[str, Any]]:
        """仅提取 frontmatter
        
        Args:
            content: 文档内容
            
        Returns:
            元数据字典或 None
        """
        metadata, _ = self.parse(content)
        return metadata if metadata else None
    
    def update(self, content: str, metadata: Dict[str, Any]) -> str:
        """更新 frontmatter
        
        Args:
            content: 文档内容
            metadata: 新的元数据
            
        Returns:
            更新后的文档内容
        """
        try:
            import yaml
            
            # 解析现有内容
            existing_meta, body = self.parse(content)
            
            # 合并元数据
            merged = {**existing_meta, **metadata}
            
            # 生成新的 frontmatter
            yaml_str = yaml.dump(
                merged,
                allow_unicode=True,
                default_flow_style=False,
                sort_keys=False
            )
            
            return f'---\n{yaml_str}---\n\n{body.lstrip()}'
            
        except Exception:
            return content
    
    def remove(self, content: str) -> str:
        """移除 frontmatter
        
        Args:
            content: 文档内容
            
        Returns:
            移除 frontmatter 后的内容
        """
        _, body = self.parse(content)
        return body.lstrip()
    
    def create_default(self, title: str, author: str = "") -> Dict[str, Any]:
        """创建默认 frontmatter
        
        Args:
            title: 文档标题
            author: 作者（可选）
            
        Returns:
            默认元数据字典
        """
        metadata = {
            'title': title,
            'created': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'modified': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'tags': [],
        }
        
        if author:
            metadata['author'] = author
        
        return metadata
    
    def validate(self, metadata: Dict[str, Any]) -> List[str]:
        """验证 frontmatter
        
        Args:
            metadata: 元数据字典
            
        Returns:
            错误列表（空列表表示验证通过）
        """
        errors = []
        
        # 检查必需字段
        # （目前没有强制必需字段）
        
        # 检查字段类型
        if 'title' in metadata and not isinstance(metadata['title'], str):
            errors.append("'title' 必须是字符串")
        
        if 'tags' in metadata:
            if not isinstance(metadata['tags'], list):
                errors.append("'tags' 必须是列表")
            elif not all(isinstance(t, str) for t in metadata['tags']):
                errors.append("'tags' 中的所有项必须是字符串")
        
        if 'created' in metadata:
            if not self._is_valid_date(metadata['created']):
                errors.append("'created' 日期格式无效")
        
        if 'modified' in metadata:
            if not self._is_valid_date(metadata['modified']):
                errors.append("'modified' 日期格式无效")
        
        return errors
    
    def _is_valid_date(self, value: Any) -> bool:
        """检查是否为有效日期
        
        Args:
            value: 要检查的值
            
        Returns:
            是否为有效日期
        """
        if isinstance(value, datetime):
            return True
        
        if isinstance(value, str):
            formats = [
                '%Y-%m-%d',
                '%Y-%m-%d %H:%M:%S',
                '%Y/%m/%d',
                '%Y/%m/%d %H:%M:%S',
            ]
            
            for fmt in formats:
                try:
                    datetime.strptime(value, fmt)
                    return True
                except ValueError:
                    continue
        
        return False
    
    def to_yaml(self, metadata: Dict[str, Any]) -> str:
        """将元数据转换为 YAML 字符串
        
        Args:
            metadata: 元数据字典
            
        Returns:
            YAML 字符串
        """
        try:
            import yaml
            return yaml.dump(
                metadata,
                allow_unicode=True,
                default_flow_style=False,
                sort_keys=False
            )
        except Exception:
            return ""
    
    def from_yaml(self, yaml_str: str) -> Optional[Dict[str, Any]]:
        """从 YAML 字符串解析元数据
        
        Args:
            yaml_str: YAML 字符串
            
        Returns:
            元数据字典或 None
        """
        try:
            import yaml
            return yaml.safe_load(yaml_str)
        except Exception:
            return None
    
    def get_field(self, content: str, field: str, default: Any = None) -> Any:
        """获取指定字段的值
        
        Args:
            content: 文档内容
            field: 字段名
            default: 默认值
            
        Returns:
            字段值或默认值
        """
        metadata = self.extract(content)
        if metadata:
            return metadata.get(field, default)
        return default
    
    def set_field(self, content: str, field: str, value: Any) -> str:
        """设置指定字段的值
        
        Args:
            content: 文档内容
            field: 字段名
            value: 字段值
            
        Returns:
            更新后的文档内容
        """
        return self.update(content, {field: value})
    
    def add_tag(self, content: str, tag: str) -> str:
        """添加标签
        
        Args:
            content: 文档内容
            tag: 标签名
            
        Returns:
            更新后的文档内容
        """
        metadata = self.extract(content) or {}
        tags = metadata.get('tags', [])
        
        if not isinstance(tags, list):
            tags = [tags] if tags else []
        
        if tag not in tags:
            tags.append(tag)
        
        return self.update(content, {'tags': tags})
    
    def remove_tag(self, content: str, tag: str) -> str:
        """移除标签
        
        Args:
            content: 文档内容
            tag: 标签名
            
        Returns:
            更新后的文档内容
        """
        metadata = self.extract(content) or {}
        tags = metadata.get('tags', [])
        
        if not isinstance(tags, list):
            tags = [tags] if tags else []
        
        if tag in tags:
            tags.remove(tag)
        
        return self.update(content, {'tags': tags})
