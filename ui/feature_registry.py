# -*- coding: utf-8 -*-
"""
Feature Registry - 统一管理 Feature 的注册和初始化

防止重复初始化，提供单例访问模式。
"""

import logging
from typing import Any, Dict, Optional, Type, Callable

logger = logging.getLogger(__name__)


class FeatureRegistry:
    """
    Feature 统一注册和初始化，防止重复创建。
    
    每个 Feature 类型在 App 生命周期内只创建一个实例。
    """
    
    def __init__(self, app):
        """
        初始化 FeatureRegistry。
        
        Args:
            app: 主应用实例
        """
        self.app = app
        self._features: Dict[str, Any] = {}
        self._initialized = False
    
    def register(self, name: str, feature_class: Type, *args, **kwargs) -> Any:
        """
        注册并初始化 Feature，如果已存在则返回现有实例。
        
        Args:
            name: Feature 的唯一标识名称
            feature_class: Feature 类
            *args: 传递给 Feature 构造函数的位置参数
            **kwargs: 传递给 Feature 构造函数的关键字参数
            
        Returns:
            Feature 实例
        """
        if name in self._features:
            logger.warning(f"Feature '{name}' already registered, returning existing instance")
            return self._features[name]
        
        try:
            instance = feature_class(self.app, *args, **kwargs)
            self._features[name] = instance
            return instance
        except Exception as e:
            logger.error(f"Failed to initialize feature '{name}': {e}")
            raise
    
    def get(self, name: str) -> Optional[Any]:
        """
        获取已注册的 Feature。
        
        Args:
            name: Feature 的唯一标识名称
            
        Returns:
            Feature 实例，如果不存在则返回 None
        """
        return self._features.get(name)
    
    def has(self, name: str) -> bool:
        """
        检查 Feature 是否已注册。
        
        Args:
            name: Feature 的唯一标识名称
            
        Returns:
            True 如果已注册，否则 False
        """
        return name in self._features
    
    def get_all(self) -> Dict[str, Any]:
        """
        获取所有已注册的 Feature。
        
        Returns:
            包含所有 Feature 的字典
        """
        return self._features.copy()
    
    def count(self) -> int:
        """
        获取已注册的 Feature 数量。
        
        Returns:
            Feature 数量
        """
        return len(self._features)
    
    def initialize_all(self, feature_definitions: list) -> None:
        """
        批量初始化所有 Feature。
        
        Args:
            feature_definitions: Feature 定义列表，每个元素为 (name, class, args, kwargs) 元组
        """
        if self._initialized:
            logger.warning("Features already initialized, skipping")
            return
        
        for definition in feature_definitions:
            if len(definition) == 2:
                name, feature_class = definition
                args, kwargs = (), {}
            elif len(definition) == 3:
                name, feature_class, args = definition
                kwargs = {}
            else:
                name, feature_class, args, kwargs = definition
            
            try:
                self.register(name, feature_class, *args, **kwargs)
            except Exception as e:
                logger.error(f"Failed to initialize feature '{name}': {e}")
        
        self._initialized = True
    
    @property
    def is_initialized(self) -> bool:
        """检查是否已完成批量初始化。"""
        return self._initialized
