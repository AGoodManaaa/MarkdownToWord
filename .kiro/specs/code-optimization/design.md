# Design Document: Code Optimization

## Overview

本设计文档描述了 MarkdowntoWord 项目的代码优化方案，主要包括：
1. 代码结构重构 - 拆分 `gui.py` 中的 `App` 类，消除重复初始化
2. 性能优化 - 改进实时预览性能，实现异步图片下载和缓存

## Architecture

### 当前架构问题

```
gui.py (1288行)
└── App 类
    ├── __init__() - 初始化 30+ 个 Feature（有重复）
    ├── _init_ui() - UI 初始化 + 快捷键绑定
    ├── _create_header() - 200+ 行 Header 创建代码
    ├── _create_main_content() - 100+ 行主内容创建
    ├── _create_preview_panel() - 预览面板创建
    └── 50+ 个方法混杂在一起
```

### 目标架构

```
gui.py (精简后 ~400行)
└── App 类
    ├── __init__() - 委托给 FeatureRegistry
    ├── _init_ui() - 委托给 UIBuilder
    └── 核心业务方法

ui/
├── app_builder.py (新增)
│   ├── UIBuilder - 统一 UI 构建入口
│   ├── HeaderBuilder - Header 构建
│   ├── MainContentBuilder - 主内容构建
│   └── PreviewPanelBuilder - 预览面板构建
├── keyboard_shortcuts.py (新增)
│   └── KeyboardShortcutsManager - 快捷键管理
└── feature_registry.py (新增)
    └── FeatureRegistry - Feature 统一注册和初始化

utils.py
├── ImageDownloader (新增) - 异步图片下载
└── ImageCache (新增) - 图片缓存管理
```

## Components and Interfaces

### 1. FeatureRegistry - Feature 统一注册

```python
class FeatureRegistry:
    """Feature 统一注册和初始化，防止重复创建"""
    
    def __init__(self, app):
        self.app = app
        self._features: Dict[str, Any] = {}
        self._initialized = False
    
    def register(self, name: str, feature_class: type, *args, **kwargs) -> Any:
        """注册并初始化 Feature，如果已存在则返回现有实例"""
        if name in self._features:
            import logging
            logging.warning(f"Feature '{name}' already registered, returning existing instance")
            return self._features[name]
        
        instance = feature_class(self.app, *args, **kwargs)
        self._features[name] = instance
        return instance
    
    def get(self, name: str) -> Any:
        """获取已注册的 Feature"""
        return self._features.get(name)
    
    def initialize_all(self):
        """批量初始化所有 Feature"""
        if self._initialized:
            return
        
        # 按依赖顺序初始化
        self._register_core_features()
        self._register_phase1_features()
        self._register_phase3_features()
        self._register_phase4_features()
        
        self._initialized = True
```

### 2. UIBuilder - UI 构建器

```python
class UIBuilder:
    """统一 UI 构建入口"""
    
    def __init__(self, app):
        self.app = app
        self.header_builder = HeaderBuilder(app)
        self.main_content_builder = MainContentBuilder(app)
    
    def build(self):
        """构建完整 UI"""
        self.header_builder.build()
        self.app.status_bar_feature.create()
        self.main_content_builder.build()


class HeaderBuilder:
    """Header 构建器"""
    
    def __init__(self, app):
        self.app = app
    
    def build(self):
        """构建 Header"""
        self._create_frame()
        self._create_logo()
        self._create_toolbar()
        self._create_right_buttons()
    
    def _create_toolbar(self):
        """创建工具栏按钮"""
        tools = self._get_tool_definitions()
        for icon, tip, cmd, shortcut in tools:
            self._create_tool_button(icon, tip, cmd, shortcut)
    
    def _get_tool_definitions(self) -> List[Tuple]:
        """获取工具按钮定义（可配置）"""
        return [
            ("📂", "打开", self.app.open_file, "Ctrl+O"),
            ("💾", "保存", self.app.save_file, "Ctrl+S"),
            # ... 其他按钮
        ]
```

### 3. KeyboardShortcutsManager - 快捷键管理

```python
class KeyboardShortcutsManager:
    """快捷键统一管理"""
    
    DEFAULT_SHORTCUTS = {
        '<Control-o>': 'open_file',
        '<Control-s>': 'save_file',
        '<Control-Shift-s>': 'export_to_word',
        '<Control-f>': 'show_search_dialog',
        '<Control-z>': '_undo',
        '<Control-y>': '_redo',
        '<Control-p>': 'toggle_preview',
        '<F1>': 'show_help',
        '<F11>': 'focus_mode.toggle',
        '<F12>': 'reading_mode.toggle',
        # ... 其他快捷键
    }
    
    def __init__(self, app):
        self.app = app
        self.shortcuts = {}
    
    def load_from_config(self, config: dict = None):
        """从配置加载快捷键（支持自定义）"""
        self.shortcuts = self.DEFAULT_SHORTCUTS.copy()
        if config and 'shortcuts' in config:
            self.shortcuts.update(config['shortcuts'])
    
    def bind_all(self):
        """绑定所有快捷键"""
        for key, action in self.shortcuts.items():
            self._bind_shortcut(key, action)
    
    def _bind_shortcut(self, key: str, action: str):
        """绑定单个快捷键"""
        handler = self._resolve_action(action)
        if handler:
            self.app.bind(key, lambda e, h=handler: h())
    
    def _resolve_action(self, action: str) -> Callable:
        """解析 action 字符串为可调用对象"""
        parts = action.split('.')
        obj = self.app
        for part in parts:
            obj = getattr(obj, part, None)
            if obj is None:
                return None
        return obj if callable(obj) else None
```

### 4. ImageDownloader - 异步图片下载

```python
from concurrent.futures import ThreadPoolExecutor
from typing import Callable, Optional
import threading

class ImageDownloader:
    """异步图片下载器"""
    
    def __init__(self, max_workers: int = 4, cache: 'ImageCache' = None):
        self.executor = ThreadPoolExecutor(max_workers=max_workers)
        self.cache = cache
        self._pending: Dict[str, Future] = {}
        self._lock = threading.Lock()
    
    def download_async(
        self, 
        url: str, 
        on_complete: Callable[[str, Optional[str]], None],
        on_error: Callable[[str, Exception], None] = None
    ) -> None:
        """
        异步下载图片
        
        Args:
            url: 图片 URL
            on_complete: 下载完成回调 (url, local_path)
            on_error: 下载失败回调 (url, exception)
        """
        # 检查缓存
        if self.cache:
            cached_path = self.cache.get(url)
            if cached_path:
                on_complete(url, cached_path)
                return
        
        # 检查是否已在下载中
        with self._lock:
            if url in self._pending:
                return  # 避免重复下载
            
            future = self.executor.submit(self._download, url)
            self._pending[url] = future
        
        def callback(f):
            with self._lock:
                self._pending.pop(url, None)
            
            try:
                local_path = f.result()
                if local_path and self.cache:
                    self.cache.put(url, local_path)
                on_complete(url, local_path)
            except Exception as e:
                if on_error:
                    on_error(url, e)
                else:
                    on_complete(url, None)
        
        future.add_done_callback(callback)
    
    def _download(self, url: str) -> Optional[str]:
        """实际下载逻辑（在线程中执行）"""
        import requests
        import tempfile
        import os
        
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        
        ext = self._get_extension(url, response.headers.get('Content-Type', ''))
        fd, temp_path = tempfile.mkstemp(suffix=ext)
        os.close(fd)
        
        with open(temp_path, 'wb') as f:
            f.write(response.content)
        
        return temp_path
    
    def shutdown(self):
        """关闭下载器"""
        self.executor.shutdown(wait=False)
```

### 5. ImageCache - 图片缓存

```python
import os
import json
import hashlib
from collections import OrderedDict
from typing import Optional

class ImageCache:
    """图片缓存管理（LRU 策略）"""
    
    def __init__(self, cache_dir: str = None, max_size_mb: int = 100):
        self.cache_dir = cache_dir or os.path.join(
            os.path.expanduser('~'), '.md2word_cache', 'images'
        )
        self.max_size_bytes = max_size_mb * 1024 * 1024
        self.index_file = os.path.join(self.cache_dir, 'index.json')
        self._index: OrderedDict[str, dict] = OrderedDict()
        self._current_size = 0
        
        self._ensure_cache_dir()
        self._load_index()
    
    def get(self, url: str) -> Optional[str]:
        """获取缓存的图片路径"""
        key = self._url_to_key(url)
        if key not in self._index:
            return None
        
        entry = self._index[key]
        local_path = entry['path']
        
        if not os.path.exists(local_path):
            del self._index[key]
            self._save_index()
            return None
        
        # LRU: 移到末尾
        self._index.move_to_end(key)
        return local_path
    
    def put(self, url: str, local_path: str) -> None:
        """添加图片到缓存"""
        key = self._url_to_key(url)
        
        # 复制到缓存目录
        ext = os.path.splitext(local_path)[1]
        cache_path = os.path.join(self.cache_dir, f"{key}{ext}")
        
        import shutil
        shutil.copy2(local_path, cache_path)
        
        file_size = os.path.getsize(cache_path)
        
        self._index[key] = {
            'url': url,
            'path': cache_path,
            'size': file_size
        }
        self._current_size += file_size
        
        # 检查是否需要清理
        self._evict_if_needed()
        self._save_index()
    
    def _evict_if_needed(self):
        """LRU 清理"""
        while self._current_size > self.max_size_bytes and self._index:
            key, entry = self._index.popitem(last=False)
            try:
                os.remove(entry['path'])
            except OSError:
                pass
            self._current_size -= entry.get('size', 0)
    
    def _url_to_key(self, url: str) -> str:
        """URL 转缓存 key"""
        return hashlib.md5(url.encode()).hexdigest()
    
    def _load_index(self):
        """加载缓存索引"""
        if os.path.exists(self.index_file):
            try:
                with open(self.index_file, 'r') as f:
                    data = json.load(f)
                    self._index = OrderedDict(data.get('entries', []))
                    self._current_size = sum(
                        e.get('size', 0) for e in self._index.values()
                    )
            except (json.JSONDecodeError, IOError):
                pass
    
    def _save_index(self):
        """保存缓存索引"""
        try:
            with open(self.index_file, 'w') as f:
                json.dump({'entries': list(self._index.items())}, f)
        except IOError:
            pass
```

### 6. 增量预览渲染

```python
class IncrementalPreviewRenderer:
    """增量预览渲染器"""
    
    def __init__(self, preview_widget):
        self.preview = preview_widget
        self._last_content = ""
        self._last_blocks = []
    
    def render(self, content: str) -> None:
        """增量渲染"""
        if not content:
            self._full_render("")
            return
        
        # 计算差异
        new_blocks = self._parse_blocks(content)
        diff = self._compute_diff(self._last_blocks, new_blocks)
        
        if diff['type'] == 'full':
            self._full_render(content)
        else:
            self._incremental_render(diff)
        
        self._last_content = content
        self._last_blocks = new_blocks
    
    def _compute_diff(self, old_blocks, new_blocks) -> dict:
        """计算块级差异"""
        # 如果变化超过 30%，使用全量渲染
        if len(old_blocks) == 0 or len(new_blocks) == 0:
            return {'type': 'full'}
        
        # 简单的行级比较
        changed_ratio = self._calculate_change_ratio(old_blocks, new_blocks)
        if changed_ratio > 0.3:
            return {'type': 'full'}
        
        # 找出变化的块
        changes = []
        for i, (old, new) in enumerate(zip(old_blocks, new_blocks)):
            if old != new:
                changes.append({'index': i, 'old': old, 'new': new})
        
        # 处理新增/删除的块
        if len(new_blocks) > len(old_blocks):
            for i in range(len(old_blocks), len(new_blocks)):
                changes.append({'index': i, 'old': None, 'new': new_blocks[i]})
        elif len(new_blocks) < len(old_blocks):
            for i in range(len(new_blocks), len(old_blocks)):
                changes.append({'index': i, 'old': old_blocks[i], 'new': None})
        
        return {'type': 'incremental', 'changes': changes}
    
    def _incremental_render(self, diff: dict):
        """增量更新"""
        for change in diff['changes']:
            idx = change['index']
            if change['new'] is None:
                # 删除块
                self._remove_block(idx)
            elif change['old'] is None:
                # 新增块
                self._insert_block(idx, change['new'])
            else:
                # 更新块
                self._update_block(idx, change['new'])
```

## Data Models

### 配置数据结构

```python
# 快捷键配置
shortcuts_config = {
    'shortcuts': {
        '<Control-o>': 'open_file',
        '<Control-s>': 'save_file',
        # 用户可自定义覆盖
    }
}

# 图片缓存配置
cache_config = {
    'image_cache': {
        'enabled': True,
        'max_size_mb': 100,
        'cache_dir': None  # None 表示使用默认目录
    }
}

# 预览性能配置
preview_config = {
    'preview': {
        'debounce_ms': 300,
        'throttle_ms': 120,
        'incremental_threshold': 0.3,  # 变化超过 30% 使用全量渲染
        'max_render_time_ms': 200
    }
}
```

## Correctness Properties

*A property is a characteristic or behavior that should hold true across all valid executions of a system-essentially, a formal statement about what the system should do. Properties serve as the bridge between human-readable specifications and machine-verifiable correctness guarantees.*

### Property 1: Feature 单例性
*For any* Feature 类型，在 App 生命周期内，FeatureRegistry 应该只创建一个实例
**Validates: Requirements 1.1**

### Property 2: 快捷键配置一致性
*For any* 快捷键配置，加载后绑定的快捷键应该与配置完全一致
**Validates: Requirements 3.3**

### Property 3: 预览渲染性能
*For any* 大于 10KB 的文档，预览渲染时间应该小于 200ms
**Validates: Requirements 4.1**

### Property 4: 防抖有效性
*For any* 连续的快速输入序列，只有最后一次输入触发预览更新
**Validates: Requirements 4.2**

### Property 5: 增量渲染正确性
*For any* 文档变化，增量渲染后的结果应该与全量渲染结果一致
**Validates: Requirements 4.3**

### Property 6: 异步下载非阻塞
*For any* 图片下载请求，download_async 应该立即返回，不阻塞调用线程
**Validates: Requirements 5.1, 5.3**

### Property 7: 并发下载限制
*For any* 并发下载请求数量，活跃线程数不应超过 max_workers
**Validates: Requirements 5.4**

### Property 8: 缓存命中
*For any* 已缓存的图片 URL，第二次请求应该直接返回缓存路径，不发起网络请求
**Validates: Requirements 6.2**

### Property 9: LRU 缓存淘汰
*For any* 缓存大小超过限制时，最久未使用的条目应该被移除
**Validates: Requirements 6.3**

### Property 10: 缓存持久化
*For any* 缓存条目，应用重启后应该能够恢复
**Validates: Requirements 6.4**

## Error Handling

### 图片下载错误处理

```python
def handle_image_download_error(url: str, error: Exception):
    """图片下载失败处理"""
    import logging
    logging.warning(f"Image download failed: {url}, error: {error}")
    
    # 返回错误占位符路径或 None
    # UI 层负责显示错误占位符
```

### Feature 初始化错误处理

```python
def safe_register_feature(registry, name, feature_class, *args, **kwargs):
    """安全注册 Feature，捕获初始化错误"""
    try:
        return registry.register(name, feature_class, *args, **kwargs)
    except Exception as e:
        import logging
        logging.error(f"Failed to initialize feature '{name}': {e}")
        return None
```

## Testing Strategy

### 双重测试方法

本项目采用单元测试和属性测试相结合的方式：

1. **单元测试** - 验证具体示例和边界情况
2. **属性测试** - 验证通用属性在所有输入上成立

### 属性测试框架

使用 `hypothesis` 库进行属性测试：

```python
from hypothesis import given, strategies as st

@given(st.lists(st.text()))
def test_feature_singleton_property(feature_names):
    """Property 1: Feature 单例性"""
    # 测试实现
    pass
```

### 测试覆盖范围

| 组件 | 单元测试 | 属性测试 |
|------|---------|---------|
| FeatureRegistry | ✓ | ✓ (Property 1) |
| KeyboardShortcutsManager | ✓ | ✓ (Property 2) |
| IncrementalPreviewRenderer | ✓ | ✓ (Property 3, 4, 5) |
| ImageDownloader | ✓ | ✓ (Property 6, 7) |
| ImageCache | ✓ | ✓ (Property 8, 9, 10) |

