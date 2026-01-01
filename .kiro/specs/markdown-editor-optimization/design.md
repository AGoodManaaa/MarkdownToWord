# Design Document: Markdown Editor Optimization

## Overview

本设计文档描述了将现有Markdown编辑器优化为市面上最佳编辑器的技术方案。基于对现有代码库的深入分析，我们识别出以下关键问题和改进机会：

### 已识别的Bug和问题

1. **撤销/重做系统不稳定** - `UndoRedoFeature`的`enabled`默认为`False`，导致撤销功能经常失效
2. **预览同步不精确** - 滚动同步基于简单的位置比例，对于不同高度的元素会产生偏差
3. **大文档性能问题** - 每次按键都触发全量语法高亮和预览渲染
4. **主题切换闪烁** - 切换主题时没有过渡动画，体验生硬
5. **工具栏按钮过多** - 18个工具按钮挤在一起，视觉混乱
6. **快捷键冲突** - `Ctrl+Shift+C`同时绑定了"复制"和"协作"功能

### 改进目标

- 性能：大文档（10000+行）流畅编辑
- 稳定性：撤销/重做100%可靠
- 体验：现代化UI，流畅动画
- 功能：智能编辑辅助

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                        App (gui.py)                         │
├─────────────────────────────────────────────────────────────┤
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────────────┐  │
│  │   Header    │  │  StatusBar  │  │    MainContainer    │  │
│  │  (Toolbar)  │  │             │  │  ┌───────┬───────┐  │  │
│  └─────────────┘  └─────────────┘  │  │Sidebar│ Main  │  │  │
│                                     │  │       │ Frame │  │  │
│  ┌─────────────────────────────┐   │  │       ├───────┤  │  │
│  │      Feature Modules        │   │  │       │Editor │  │  │
│  │  ┌─────┐ ┌─────┐ ┌─────┐   │   │  │       │Preview│  │  │
│  │  │Undo │ │Sync │ │Theme│   │   │  └───────┴───────┘  │  │
│  │  │Redo │ │     │ │     │   │   └─────────────────────┘  │
│  │  └─────┘ └─────┘ └─────┘   │                            │
│  └─────────────────────────────┘                            │
├─────────────────────────────────────────────────────────────┤
│                    Core Modules                              │
│  ┌──────────┐  ┌──────────┐  ┌──────────┐  ┌──────────┐    │
│  │ Parser   │  │Converter │  │ Handlers │  │  Styles  │    │
│  └──────────┘  └──────────┘  └──────────┘  └──────────┘    │
└─────────────────────────────────────────────────────────────┘
```

### 优化后的架构改进

1. **虚拟化渲染层** - 只渲染可视区域
2. **增量更新引擎** - 只更新变化的部分
3. **事件节流系统** - 统一管理防抖和节流
4. **主题过渡系统** - 平滑的主题切换动画

## Components and Interfaces

### 1. 性能优化组件

```python
class VirtualizedRenderer:
    """虚拟化渲染器 - 只渲染可视区域"""
    
    def __init__(self, text_widget, buffer_lines: int = 50):
        self.text_widget = text_widget
        self.buffer_lines = buffer_lines  # 上下缓冲行数
        self.rendered_range = (0, 0)  # 已渲染的行范围
    
    def get_visible_range(self) -> Tuple[int, int]:
        """获取可视区域的行范围"""
        pass
    
    def render_visible(self, content: str) -> None:
        """只渲染可视区域"""
        pass
    
    def on_scroll(self, position: float) -> None:
        """滚动时按需渲染"""
        pass


class IncrementalUpdater:
    """增量更新器 - 只更新变化的部分"""
    
    def __init__(self):
        self.last_content = ""
        self.block_cache = {}  # 块级元素缓存
    
    def diff(self, new_content: str) -> List[Change]:
        """计算内容差异"""
        pass
    
    def apply_changes(self, changes: List[Change]) -> None:
        """应用增量更新"""
        pass
```

### 2. 撤销/重做修复

```python
class ReliableUndoManager:
    """可靠的撤销管理器"""
    
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.undo_stack = []
        self.redo_stack = []
        self.enabled = True  # 默认启用！
        self.batch_mode = False  # 批量操作模式
        self.current_batch = []
    
    def begin_batch(self) -> None:
        """开始批量操作（如粘贴多行）"""
        self.batch_mode = True
        self.current_batch = []
    
    def end_batch(self) -> None:
        """结束批量操作，合并为一个撤销点"""
        if self.current_batch:
            self.undo_stack.append(BatchOperation(self.current_batch))
        self.batch_mode = False
        self.current_batch = []
```

### 3. 精确滚动同步

```python
class PreciseScrollSync:
    """精确滚动同步 - 基于行映射"""
    
    def __init__(self, editor, preview):
        self.editor = editor
        self.preview = preview
        self.line_map = {}  # 源码行 -> 预览位置映射
    
    def build_line_map(self, content: str) -> None:
        """构建行映射表"""
        pass
    
    def sync_editor_to_preview(self, editor_line: int) -> None:
        """编辑器滚动时同步预览"""
        preview_pos = self.line_map.get(editor_line)
        if preview_pos:
            self.preview.scroll_to(preview_pos)
    
    def sync_preview_to_editor(self, preview_pos: float) -> None:
        """预览滚动时同步编辑器"""
        # 反向查找最近的源码行
        pass
```

### 4. 主题过渡系统

```python
class ThemeTransition:
    """主题过渡动画"""
    
    def __init__(self, app, duration_ms: int = 300):
        self.app = app
        self.duration_ms = duration_ms
        self.steps = 10
    
    def transition_to(self, target_theme: str) -> None:
        """平滑过渡到目标主题"""
        current_colors = self._get_current_colors()
        target_colors = COLORS_DARK if target_theme == 'dark' else COLORS_LIGHT
        
        for step in range(self.steps):
            progress = step / self.steps
            interpolated = self._interpolate_colors(
                current_colors, target_colors, progress
            )
            self._apply_colors(interpolated)
            self.app.update()
            time.sleep(self.duration_ms / self.steps / 1000)
```

## Data Models

### 文档状态模型

```python
@dataclass
class DocumentState:
    """文档状态"""
    content: str
    cursor_position: str
    scroll_position: float
    selection: Optional[Tuple[str, str]]
    modified: bool
    file_path: Optional[str]
    last_saved_content: str
    undo_stack: List[Operation]
    redo_stack: List[Operation]


@dataclass
class Operation:
    """编辑操作"""
    type: str  # 'insert', 'delete', 'replace'
    position: str
    content: str
    timestamp: float
    
    def inverse(self) -> 'Operation':
        """返回逆操作"""
        if self.type == 'insert':
            return Operation('delete', self.position, self.content, time.time())
        elif self.type == 'delete':
            return Operation('insert', self.position, self.content, time.time())
```

### 渲染缓存模型

```python
@dataclass
class RenderCache:
    """渲染缓存"""
    block_id: str
    source_lines: Tuple[int, int]  # 源码行范围
    rendered_html: str
    preview_height: int
    last_modified: float


@dataclass
class LineMapping:
    """行映射"""
    source_line: int
    preview_element_id: str
    preview_y_position: int
```

## Correctness Properties

*A property is a characteristic or behavior that should hold true across all valid executions of a system-essentially, a formal statement about what the system should do. Properties serve as the bridge between human-readable specifications and machine-verifiable correctness guarantees.*

### Property 1: 撤销/重做一致性
*For any* 编辑操作序列，执行撤销后再执行重做，文档内容应与撤销前完全一致
**Validates: Requirements 2.1, 6.1**

### Property 2: 滚动同步精确性
*For any* 编辑器滚动位置，预览区对应位置的源码行号与编辑器当前行号误差不超过1行
**Validates: Requirements 3.1**

### Property 3: 增量渲染正确性
*For any* 文档修改，增量渲染的结果应与全量渲染的结果完全一致
**Validates: Requirements 1.3, 3.4**

### Property 4: 格式化操作可逆性
*For any* 选中文本，应用格式（如加粗）后再次应用相同格式，应移除格式标记
**Validates: Requirements 2.1**

### Property 5: 导出格式保真性
*For any* Markdown文档，导出为Word后再转回Markdown，核心内容结构应保持不变
**Validates: Requirements 4.1, 4.2, 4.3, 4.4**

### Property 6: 搜索结果完整性
*For any* 搜索词，高亮的匹配数量应等于文档中实际出现的次数
**Validates: Requirements 7.2, 7.4**

### Property 7: 标签页状态一致性
*For any* 打开的文件，标签页显示的修改状态应与实际内容是否修改一致
**Validates: Requirements 9.5**

### Property 8: 主题颜色对比度
*For any* 深色模式下的文本颜色和背景颜色组合，对比度应不低于4.5:1（WCAG AA标准）
**Validates: Requirements 5.5**

### Property 9: 快捷键唯一性
*For any* 已注册的快捷键，不应存在重复绑定到不同功能的情况
**Validates: Requirements 8.1, 8.3**

### Property 10: 自动保存完整性
*For any* 自动保存的内容，恢复后应与保存时的内容完全一致
**Validates: Requirements 6.1, 6.5**

## Error Handling

### 错误分类和处理策略

| 错误类型 | 处理策略 | 用户提示 |
|---------|---------|---------|
| 文件读取失败 | 显示错误对话框，提供重试选项 | "无法打开文件: {原因}" |
| 文件保存失败 | 自动备份到临时目录，提示用户 | "保存失败，已备份到: {路径}" |
| 导出失败 | 记录详细日志，显示错误位置 | "导出失败（第{行}行）: {原因}" |
| 渲染超时 | 显示加载指示器，允许取消 | "正在渲染大文档..." |
| 插件错误 | 隔离错误，禁用问题插件 | "插件{名称}已禁用: {原因}" |
| 网络错误 | 自动重试，离线模式降级 | "网络连接中断，已切换到离线模式" |

### 错误恢复机制

```python
class ErrorRecovery:
    """错误恢复机制"""
    
    def __init__(self, app):
        self.app = app
        self.backup_dir = Path.home() / '.md2word_backup'
        self.backup_dir.mkdir(exist_ok=True)
    
    def create_backup(self, content: str, filename: str) -> Path:
        """创建备份"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_path = self.backup_dir / f"{filename}_{timestamp}.md"
        backup_path.write_text(content, encoding='utf-8')
        return backup_path
    
    def recover_from_crash(self) -> Optional[str]:
        """从崩溃中恢复"""
        # 查找最近的自动保存文件
        auto_save_file = self.backup_dir / '.autosave.md'
        if auto_save_file.exists():
            return auto_save_file.read_text(encoding='utf-8')
        return None
```

## Testing Strategy

### 双重测试方法

本项目采用单元测试和属性测试相结合的方法：

- **单元测试**: 验证具体示例和边界情况
- **属性测试**: 验证应在所有输入上成立的通用属性

### 属性测试框架

使用 **Hypothesis** 作为Python属性测试库。

### 测试配置

每个属性测试运行最少100次迭代。

### 测试文件结构

```
tests/
├── unit/
│   ├── test_parser.py
│   ├── test_converter.py
│   ├── test_undo_redo.py
│   └── test_scroll_sync.py
├── property/
│   ├── test_undo_redo_properties.py
│   ├── test_scroll_sync_properties.py
│   ├── test_render_properties.py
│   └── test_format_properties.py
└── integration/
    ├── test_export_workflow.py
    └── test_edit_workflow.py
```

### 属性测试示例

```python
from hypothesis import given, strategies as st, settings

class TestUndoRedoProperties:
    """撤销/重做属性测试"""
    
    @given(st.lists(st.text(min_size=1, max_size=100), min_size=1, max_size=20))
    @settings(max_examples=100)
    def test_undo_redo_roundtrip(self, operations):
        """
        **Feature: markdown-editor-optimization, Property 1: 撤销/重做一致性**
        **Validates: Requirements 2.1, 6.1**
        
        对于任意编辑操作序列，撤销后重做应恢复原状态
        """
        editor = MockEditor()
        
        # 执行所有操作
        for op in operations:
            editor.insert(op)
        
        content_before_undo = editor.get_content()
        
        # 撤销所有操作
        for _ in operations:
            editor.undo()
        
        # 重做所有操作
        for _ in operations:
            editor.redo()
        
        content_after_redo = editor.get_content()
        
        assert content_before_undo == content_after_redo


class TestScrollSyncProperties:
    """滚动同步属性测试"""
    
    @given(st.floats(min_value=0.0, max_value=1.0))
    @settings(max_examples=100)
    def test_scroll_sync_precision(self, scroll_position):
        """
        **Feature: markdown-editor-optimization, Property 2: 滚动同步精确性**
        **Validates: Requirements 3.1**
        
        对于任意滚动位置，同步误差不超过1行
        """
        sync = PreciseScrollSync(mock_editor, mock_preview)
        
        editor_line = sync.position_to_line(scroll_position)
        sync.sync_editor_to_preview(editor_line)
        preview_line = sync.get_preview_current_line()
        
        assert abs(editor_line - preview_line) <= 1
```
