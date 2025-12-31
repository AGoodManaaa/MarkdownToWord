# -*- coding: utf-8 -*-
"""知识图谱视图模块"""

from dataclasses import dataclass, field
from typing import List, Optional, Dict, Tuple
import json
import math
import random

from .links import LinkManager
from .vault import DocumentInfo


@dataclass
class GraphNode:
    """图节点"""
    id: str
    label: str
    x: float = 0.0
    y: float = 0.0
    size: float = 10.0
    color: str = "#4ECDC4"
    tags: List[str] = field(default_factory=list)
    
    def to_dict(self) -> dict:
        """转换为字典"""
        return {
            'id': self.id,
            'label': self.label,
            'x': self.x,
            'y': self.y,
            'size': self.size,
            'color': self.color,
            'tags': self.tags
        }


@dataclass
class GraphEdge:
    """图边"""
    source: str
    target: str
    weight: float = 1.0
    
    def to_dict(self) -> dict:
        """转换为字典"""
        return {
            'source': self.source,
            'target': self.target,
            'weight': self.weight
        }


@dataclass
class GraphData:
    """图数据"""
    nodes: List[GraphNode] = field(default_factory=list)
    edges: List[GraphEdge] = field(default_factory=list)
    
    @property
    def node_count(self) -> int:
        """节点数量"""
        return len(self.nodes)
    
    @property
    def edge_count(self) -> int:
        """边数量"""
        return len(self.edges)
    
    def get_node(self, node_id: str) -> Optional[GraphNode]:
        """获取节点"""
        for node in self.nodes:
            if node.id == node_id:
                return node
        return None
    
    def to_dict(self) -> dict:
        """转换为字典"""
        return {
            'nodes': [n.to_dict() for n in self.nodes],
            'edges': [e.to_dict() for e in self.edges]
        }
    
    def to_json(self) -> str:
        """转换为 JSON"""
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=2)


class GraphView:
    """知识图谱视图"""
    
    NODE_COLORS = [
        "#FF6B6B", "#4ECDC4", "#45B7D1", "#96CEB4",
        "#FFEAA7", "#DDA0DD", "#98D8C8", "#F7DC6F",
        "#BB8FCE", "#85C1E9", "#F8B500", "#00CED1"
    ]
    
    def __init__(self, link_manager: LinkManager):
        """初始化图谱视图
        
        Args:
            link_manager: 链接管理器
        """
        self.link_manager = link_manager
        self._layout_algorithm = "force_directed"
    
    def build_graph(self, documents: List[DocumentInfo]) -> GraphData:
        """构建图数据
        
        Args:
            documents: 文档列表
            
        Returns:
            GraphData 图数据
        """
        nodes = []
        edges = []
        node_ids = set()
        
        # 创建节点
        for i, doc in enumerate(documents):
            node_id = str(doc.path)
            node_ids.add(node_id)
            
            # 根据链接数量计算节点大小
            link_count = len(doc.links) + len(doc.backlinks)
            size = 10 + min(link_count * 2, 30)
            
            # 根据标签分配颜色
            color_index = hash(doc.tags[0]) % len(self.NODE_COLORS) if doc.tags else i % len(self.NODE_COLORS)
            
            nodes.append(GraphNode(
                id=node_id,
                label=doc.title or doc.filename,
                size=size,
                color=self.NODE_COLORS[color_index],
                tags=doc.tags
            ))
        
        # 创建边
        for doc in documents:
            source_id = str(doc.path)
            
            for link in doc.links:
                # 尝试解析链接目标
                target_id = self._resolve_link_target(link, node_ids)
                
                if target_id and target_id in node_ids:
                    edges.append(GraphEdge(
                        source=source_id,
                        target=target_id,
                        weight=1.0
                    ))
        
        graph = GraphData(nodes=nodes, edges=edges)
        
        # 计算布局
        graph = self.calculate_layout(graph)
        
        return graph
    
    def _resolve_link_target(self, link: str, node_ids: set) -> Optional[str]:
        """解析链接目标
        
        Args:
            link: 链接文本
            node_ids: 已知节点 ID 集合
            
        Returns:
            目标节点 ID 或 None
        """
        # 添加 .md 扩展名
        if not link.endswith('.md'):
            link_with_ext = f'{link}.md'
        else:
            link_with_ext = link
        
        # 直接匹配
        if link_with_ext in node_ids:
            return link_with_ext
        
        # 尝试匹配文件名
        for node_id in node_ids:
            if node_id.endswith(f'/{link_with_ext}') or node_id == link_with_ext:
                return node_id
        
        return None
    
    def filter_by_tag(self, graph: GraphData, tag: str) -> GraphData:
        """按标签过滤图
        
        Args:
            graph: 原始图数据
            tag: 标签名
            
        Returns:
            过滤后的 GraphData
        """
        # 过滤节点
        filtered_nodes = [
            node for node in graph.nodes
            if tag in node.tags or any(t.startswith(f'{tag}/') for t in node.tags)
        ]
        
        filtered_node_ids = {node.id for node in filtered_nodes}
        
        # 过滤边
        filtered_edges = [
            edge for edge in graph.edges
            if edge.source in filtered_node_ids and edge.target in filtered_node_ids
        ]
        
        return GraphData(nodes=filtered_nodes, edges=filtered_edges)
    
    def filter_by_depth(self, graph: GraphData, center: str, depth: int) -> GraphData:
        """按深度过滤（以某节点为中心）
        
        Args:
            graph: 原始图数据
            center: 中心节点 ID
            depth: 深度
            
        Returns:
            过滤后的 GraphData
        """
        if depth <= 0:
            node = graph.get_node(center)
            if node:
                return GraphData(nodes=[node], edges=[])
            return GraphData()
        
        # BFS 查找指定深度内的节点
        visited = {center}
        current_level = {center}
        
        # 构建邻接表
        adjacency: Dict[str, set] = {}
        for edge in graph.edges:
            if edge.source not in adjacency:
                adjacency[edge.source] = set()
            if edge.target not in adjacency:
                adjacency[edge.target] = set()
            adjacency[edge.source].add(edge.target)
            adjacency[edge.target].add(edge.source)
        
        for _ in range(depth):
            next_level = set()
            for node_id in current_level:
                if node_id in adjacency:
                    for neighbor in adjacency[node_id]:
                        if neighbor not in visited:
                            visited.add(neighbor)
                            next_level.add(neighbor)
            current_level = next_level
        
        # 过滤节点和边
        filtered_nodes = [node for node in graph.nodes if node.id in visited]
        filtered_edges = [
            edge for edge in graph.edges
            if edge.source in visited and edge.target in visited
        ]
        
        return GraphData(nodes=filtered_nodes, edges=filtered_edges)
    
    def calculate_layout(self, graph: GraphData, width: float = 800, height: float = 600) -> GraphData:
        """计算节点布局
        
        使用力导向布局算法
        
        Args:
            graph: 图数据
            width: 画布宽度
            height: 画布高度
            
        Returns:
            更新位置后的 GraphData
        """
        if not graph.nodes:
            return graph
        
        # 初始化随机位置
        for node in graph.nodes:
            node.x = random.uniform(50, width - 50)
            node.y = random.uniform(50, height - 50)
        
        # 构建邻接表
        adjacency: Dict[str, List[str]] = {}
        for node in graph.nodes:
            adjacency[node.id] = []
        for edge in graph.edges:
            if edge.source in adjacency:
                adjacency[edge.source].append(edge.target)
            if edge.target in adjacency:
                adjacency[edge.target].append(edge.source)
        
        # 力导向迭代
        iterations = 100
        k = math.sqrt((width * height) / len(graph.nodes))  # 理想距离
        
        for _ in range(iterations):
            # 计算斥力
            displacements: Dict[str, Tuple[float, float]] = {}
            
            for node in graph.nodes:
                dx, dy = 0.0, 0.0
                
                for other in graph.nodes:
                    if node.id == other.id:
                        continue
                    
                    delta_x = node.x - other.x
                    delta_y = node.y - other.y
                    distance = max(math.sqrt(delta_x ** 2 + delta_y ** 2), 0.01)
                    
                    # 斥力
                    force = (k ** 2) / distance
                    dx += (delta_x / distance) * force
                    dy += (delta_y / distance) * force
                
                displacements[node.id] = (dx, dy)
            
            # 计算引力（连接的节点之间）
            for edge in graph.edges:
                source_node = graph.get_node(edge.source)
                target_node = graph.get_node(edge.target)
                
                if not source_node or not target_node:
                    continue
                
                delta_x = target_node.x - source_node.x
                delta_y = target_node.y - source_node.y
                distance = max(math.sqrt(delta_x ** 2 + delta_y ** 2), 0.01)
                
                # 引力
                force = (distance ** 2) / k
                
                dx = (delta_x / distance) * force
                dy = (delta_y / distance) * force
                
                sx, sy = displacements[source_node.id]
                tx, ty = displacements[target_node.id]
                
                displacements[source_node.id] = (sx + dx, sy + dy)
                displacements[target_node.id] = (tx - dx, ty - dy)
            
            # 应用位移
            temperature = 1.0 - (_ / iterations)  # 冷却
            
            for node in graph.nodes:
                dx, dy = displacements[node.id]
                distance = max(math.sqrt(dx ** 2 + dy ** 2), 0.01)
                
                # 限制位移
                max_disp = min(distance, temperature * 50)
                
                node.x += (dx / distance) * max_disp
                node.y += (dy / distance) * max_disp
                
                # 边界约束
                node.x = max(50, min(width - 50, node.x))
                node.y = max(50, min(height - 50, node.y))
        
        return graph
    
    def get_neighbors(self, graph: GraphData, node_id: str) -> List[str]:
        """获取相邻节点
        
        Args:
            graph: 图数据
            node_id: 节点 ID
            
        Returns:
            相邻节点 ID 列表
        """
        neighbors = set()
        
        for edge in graph.edges:
            if edge.source == node_id:
                neighbors.add(edge.target)
            elif edge.target == node_id:
                neighbors.add(edge.source)
        
        return list(neighbors)
    
    def export_to_json(self, graph: GraphData) -> str:
        """导出为 JSON
        
        Args:
            graph: 图数据
            
        Returns:
            JSON 字符串
        """
        return graph.to_json()
    
    def get_statistics(self, graph: GraphData) -> dict:
        """获取图统计信息
        
        Args:
            graph: 图数据
            
        Returns:
            统计信息字典
        """
        if not graph.nodes:
            return {
                'node_count': 0,
                'edge_count': 0,
                'avg_degree': 0,
                'density': 0,
                'isolated_nodes': 0
            }
        
        # 计算度数
        degrees: Dict[str, int] = {node.id: 0 for node in graph.nodes}
        for edge in graph.edges:
            degrees[edge.source] = degrees.get(edge.source, 0) + 1
            degrees[edge.target] = degrees.get(edge.target, 0) + 1
        
        avg_degree = sum(degrees.values()) / len(degrees) if degrees else 0
        isolated_nodes = sum(1 for d in degrees.values() if d == 0)
        
        # 计算密度
        n = len(graph.nodes)
        max_edges = n * (n - 1) / 2
        density = len(graph.edges) / max_edges if max_edges > 0 else 0
        
        return {
            'node_count': len(graph.nodes),
            'edge_count': len(graph.edges),
            'avg_degree': round(avg_degree, 2),
            'density': round(density, 4),
            'isolated_nodes': isolated_nodes
        }
