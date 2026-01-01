# Design Document: Best Markdown Editor

## Overview

本设计文档描述了将现有Markdown编辑器打造为市面上最佳编辑器的技术方案。基于需求分析和代码审查，我们将实现以下核心改进：

### 设计目标

1. **简洁高效的工具栏** - 分组设计，减少视觉混乱
2. **智能编辑体验** - 多光标、自动补全、智能缩进
3. **精确的预览同步** - 基于行映射的双向同步
4. **卓越的性能** - 虚拟滚动、增量渲染
5. **流畅的主题系统** - 平滑过渡、对比度保证
6. **统一的导出体验** - 一站式导出中心

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                         App (gui.py)                            │
├─────────────────────────────────────────────────────────────────┤
│  ┌──────────────────────────────────────────────────────────┐   │
│  │                    Toolbar (重新设计)                      │   │
│  │  [文件组] [编辑组] [视图组] [导出▼] [工具▼] [设置]        │   │
│  └──────────────────────────────────────────────────────────┘   │
│                                                                  │
│  ┌────────┐  ┌────────────────────────────────────────────┐    │
│  │Sidebar │  │              Main Content                   │    │
│  │        │  │  ┌─────────────────┬─────────────────────┐ │    │
│  │ 文件夹  │  │  │     Editor      │      Preview        │ │    │
│  │ 大纲    │  │  │  (智能编辑增强)  │  (精确同步)         │ │    │
│  │ 最近    │  │  │                 │                     │ │    │
│  │        │  │  └─────────────────┴─────────────────────┘ │    │
│  └────────┘  └────────────────────────────────────────────┘    │
│                                                                  │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │                    StatusBar (信息优化)                    │   │
│  │  [行:列] [字数] [选区] [编码] [保存状态] [任务进度]        │   │
│  └──────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────┘
```

### 核心模块架构

```
ui/
├── features/
│   ├── toolbar_redesign.py      # 工具栏重新设计
│   ├── smart_editing.py         # 智能编辑增强
│   ├── precise_scroll_sync.py   # 精确滚动同步
│   ├── virtual_renderer.py      # 虚拟化渲染
│   ├── theme_transition.py      # 主题过渡动画
│   ├── export_center.py         # 统一导出中心
│   ├── shortcut_manager.py      # 快捷键管理
│   └── status_bar_enhanced.py   # 状态栏增强
```

## Components and Interfaces

### 1. 工具栏重新设计组件

```python
class ToolbarRedesign:
    """重新设计的工具栏 - 分组+下拉菜单"""
    
    def __init__(self, app):
        self.app = app
        self.groups = {
            'file': ['new', 'open', 'save'],
            'edit': ['undo', 'redo', 'cut', 'copy'],
            'view': ['preview', 'sidebar', 'minimap'],
        }
        self.dropdowns = {
            'export': ['word', 'pdf', 'html', 'batch'],
            'tools': ['format', 'toc', 'chart', 'mindmap', 'ocr'],
        }
    
    def create_toolbar(self, parent) -> ctk.CTkFrame:
        """创建分组工具栏"""
        pass
    
    def create_dropdown(self, parent, name: str, items: List) -> ctk.CTkButton:
        """创建下拉菜单按钮"""
        pass
    
    def set_compact_mode(self, compact: bool) -> None:
        """切换紧凑模式（仅图标）"""
        pass


class ToolbarTooltip:
    """带快捷键提示的工具提示"""
    
    def __init__(self, widget, text: str, shortcut: str = None, delay_ms: int = 300):
        self.widget = widget
        self.text = text
        self.shortcut = shortcut
        self.delay_ms = delay_ms
    
    def show(self) -> None:
        """显示工具提示"""
        pass
```

### 2. 智能编辑组件

```python
class SmartEditingEnhanced:
    """增强的智能编辑功能"""
    
    # 括号配对映射
    BRACKET_PAIRS = {
        '(': ')', '[': ']', '{': '}',
        '"': '"', "'": "'", '`': '`',
        '（': '）', '【': '】', '「': '」',
        '"': '"', ''': ''',
    }
    
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.cursors = []  # 多光标位置
    
    def on_bracket_input(self, char: str) -> bool:
        """处理括号输入，自动补全配对"""
        if char in self.BRACKET_PAIRS:
            self.text_widget.insert('insert', char + self.BRACKET_PAIRS[char])
            self.text_widget.mark_set('insert', 'insert-1c')
            return True
        return False
    
    def add_cursor_at_next_match(self) -> None:
        """Ctrl+D: 选中下一个相同文本"""
        pass
    
    def indent_selection(self, increase: bool = True) -> None:
        """Tab/Shift+Tab: 增加/减少选中行缩进"""
        pass
    
    def toggle_comment(self) -> None:
        """Ctrl+/: 切换注释"""
        pass
```

### 3. 精确滚动同步组件

```python
@dataclass
class LineMapping:
    """行映射数据"""
    source_line: int
    preview_element_id: str
    preview_y_position: float


class PreciseScrollSync:
    """精确滚动同步 - 基于行映射"""
    
    def __init__(self, editor, preview):
        self.editor = editor
        self.preview = preview
        self.line_map: Dict[int, LineMapping] = {}
        self._sync_lock = False
    
    def build_line_map(self, content: str) -> None:
        """构建源码行到预览位置的映射表"""
        blocks = parse_markdown(content)
        current_line = 1
        
        for block in blocks:
            # 计算块在预览中的位置
            preview_pos = self._get_preview_position(block)
            self.line_map[current_line] = LineMapping(
                source_line=current_line,
                preview_element_id=f"block_{current_line}",
                preview_y_position=preview_pos
            )
            current_line += block.content.count('\n') + 1
    
    def sync_editor_to_preview(self, editor_line: int) -> None:
        """编辑器滚动时同步预览"""
        if self._sync_lock:
            return
        
        self._sync_lock = True
        try:
            # 找到最近的映射行
            mapped_line = self._find_nearest_mapped_line(editor_line)
            if mapped_line:
                self.preview.scroll_to(mapped_line.preview_y_position)
        finally:
            self._sync_lock = False
    
    def sync_preview_to_editor(self, preview_pos: float) -> None:
        """预览滚动时同步编辑器"""
        if self._sync_lock:
            return
        
        self._sync_lock = True
        try:
            # 反向查找对应的源码行
            source_line = self._find_source_line(preview_pos)
            if source_line:
                self.editor.scroll_to_line(source_line)
        finally:
            self._sync_lock = False
```

### 4. 虚拟化渲染组件

```python
class VirtualRenderer:
    """虚拟化渲染器 - 只渲染可视区域"""
    
    def __init__(self, text_widget, buffer_lines: int = 50):
        self.text_widget = text_widget
        self.buffer_lines = buffer_lines
        self.rendered_range = (0, 0)
        self.content_cache: Dict[int, str] = {}
    
    def get_visible_range(self) -> Tuple[int, int]:
        """获取可视区域的行范围"""
        first_visible = self.text_widget.index("@0,0")
        last_visible = self.text_widget.index(f"@0,{self.text_widget.winfo_height()}")
        
        first_line = int(first_visible.split('.')[0])
        last_line = int(last_visible.split('.')[0])
        
        return (
            max(1, first_line - self.buffer_lines),
            last_line + self.buffer_lines
        )
    
    def render_visible(self) -> None:
        """只渲染可视区域的语法高亮"""
        start, end = self.get_visible_range()
        
        if (start, end) == self.rendered_range:
            return  # 无需重新渲染
        
        # 清除旧的高亮
        self._clear_highlights(self.rendered_range[0], self.rendered_range[1])
        
        # 应用新的高亮
        self._apply_highlights(start, end)
        
        self.rendered_range = (start, end)


class IncrementalPreviewUpdater:
    """增量预览更新器"""
    
    def __init__(self, preview):
        self.preview = preview
        self.block_cache: Dict[str, str] = {}
    
    def update(self, old_content: str, new_content: str) -> None:
        """增量更新预览"""
        old_blocks = parse_markdown(old_content)
        new_blocks = parse_markdown(new_content)
        
        # 找出变化的块
        changed_blocks = self._diff_blocks(old_blocks, new_blocks)
        
        # 只更新变化的块
        for block_id, block in changed_blocks:
            self._update_block(block_id, block)
```

### 5. 主题过渡组件

```python
class ThemeTransition:
    """主题过渡动画"""
    
    def __init__(self, app, duration_ms: int = 300, steps: int = 15):
        self.app = app
        self.duration_ms = duration_ms
        self.steps = steps
    
    def transition_to(self, target_theme: str) -> None:
        """平滑过渡到目标主题"""
        current_colors = self._get_current_colors()
        target_colors = COLORS_DARK if target_theme == 'dark' else COLORS_LIGHT
        
        step_duration = self.duration_ms / self.steps
        
        for step in range(self.steps + 1):
            progress = step / self.steps
            # 使用缓动函数使过渡更自然
            eased_progress = self._ease_in_out(progress)
            
            interpolated = self._interpolate_colors(
                current_colors, target_colors, eased_progress
            )
            self._apply_colors(interpolated)
            self.app.update()
            time.sleep(step_duration / 1000)
    
    def _ease_in_out(self, t: float) -> float:
        """缓动函数"""
        return t * t * (3 - 2 * t)
    
    def _interpolate_colors(self, c1: dict, c2: dict, t: float) -> dict:
        """颜色插值"""
        result = {}
        for key in c1:
            if key in c2:
                result[key] = self._interpolate_color(c1[key], c2[key], t)
        return result


class ContrastChecker:
    """对比度检查器"""
    
    WCAG_AA_RATIO = 4.5
    
    @staticmethod
    def calculate_contrast(fg: str, bg: str) -> float:
        """计算两个颜色的对比度"""
        fg_lum = ContrastChecker._get_luminance(fg)
        bg_lum = ContrastChecker._get_luminance(bg)
        
        lighter = max(fg_lum, bg_lum)
        darker = min(fg_lum, bg_lum)
        
        return (lighter + 0.05) / (darker + 0.05)
    
    @staticmethod
    def check_theme_contrast(theme: dict) -> List[Tuple[str, str, float]]:
        """检查主题中所有颜色组合的对比度"""
        issues = []
        text_colors = ['text_primary', 'text_secondary', 'text_muted']
        bg_colors = ['bg_light', 'bg_card', 'bg_sidebar']
        
        for text_key in text_colors:
            for bg_key in bg_colors:
                if text_key in theme and bg_key in theme:
                    ratio = ContrastChecker.calculate_contrast(
                        theme[text_key], theme[bg_key]
                    )
                    if ratio < ContrastChecker.WCAG_AA_RATIO:
                        issues.append((text_key, bg_key, ratio))
        
        return issues
```

### 6. 统一导出中心

```python
class ExportCenter:
    """统一导出中心"""
    
    FORMATS = {
        'word': {'name': 'Word文档', 'ext': '.docx', 'icon': '📄'},
        'pdf': {'name': 'PDF文档', 'ext': '.pdf', 'icon': '📕'},
        'html': {'name': 'HTML网页', 'ext': '.html', 'icon': '🌐'},
        'batch': {'name': '批量导出', 'ext': None, 'icon': '📦'},
    }
    
    def __init__(self, app):
        self.app = app
        self.dialog = None
    
    def show_export_dialog(self) -> None:
        """显示统一导出对话框"""
        self.dialog = ExportDialog(self.app, self)
        self.dialog.show()
    
    def export(self, format: str, options: dict) -> ExportResult:
        """执行导出"""
        if format == 'word':
            return self._export_word(options)
        elif format == 'pdf':
            return self._export_pdf(options)
        elif format == 'html':
            return self._export_html(options)
        elif format == 'batch':
            return self._export_batch(options)


class ExportDialog(ctk.CTkToplevel):
    """导出对话框"""
    
    def __init__(self, app, export_center):
        super().__init__(app)
        self.export_center = export_center
        self.selected_format = 'word'
        
        self.title("📤 导出文档")
        self.geometry("600x500")
        
        self._create_ui()
    
    def _create_ui(self):
        """创建界面"""
        # 左侧：格式选择
        # 右侧：格式特定选项
        pass
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
    word_count: int
    char_count: int
    line_count: int


@dataclass
class EditorStatistics:
    """编辑器统计信息"""
    word_count: int
    char_count: int
    char_count_no_spaces: int
    line_count: int
    paragraph_count: int
    selection_char_count: int
    selection_line_count: int
    cursor_line: int
    cursor_column: int
```

### 快捷键配置模型

```python
@dataclass
class ShortcutConfig:
    """快捷键配置"""
    id: str
    name: str
    description: str
    default_key: str
    current_key: str
    category: str
    
    def has_conflict(self, other: 'ShortcutConfig') -> bool:
        """检查是否与另一个快捷键冲突"""
        return self.current_key == other.current_key and self.id != other.id
```

## Correctness Properties

*A property is a characteristic or behavior that should hold true across all valid executions of a system-essentially, a formal statement about what the system should do. Properties serve as the bridge between human-readable specifications and machine-verifiable correctness guarantees.*

### Property 1: 括号自动补全一致性

*For any* 括号或引号字符输入，编辑器应自动插入配对字符，且光标应位于配对字符之间
**Validates: Requirements 2.2**

### Property 2: 多行缩进一致性

*For any* 选中的多行文本，按Tab后每行都应增加相同的缩进量（4个空格）
**Validates: Requirements 2.3, 2.4**

### Property 3: 预览同步精确性

*For any* 编辑器滚动位置，预览区对应位置的源码行号与编辑器当前行号误差不超过2行
**Validates: Requirements 3.1**

### Property 4: 增量渲染正确性

*For any* 文档修改，增量渲染的结果应与全量渲染的结果在视觉上完全一致
**Validates: Requirements 3.4**

### Property 5: 主题对比度合规性

*For any* 深色模式下的文本颜色和背景颜色组合，对比度应不低于4.5:1（WCAG AA标准）
**Validates: Requirements 5.2**

### Property 6: 主题保存往返一致性

*For any* 自定义主题，保存为JSON后再加载，应与原主题完全一致
**Validates: Requirements 5.4, 5.5**

### Property 7: 快捷键唯一性

*For any* 已注册的快捷键配置，不应存在两个不同功能绑定到相同快捷键的情况
**Validates: Requirements 7.2, 7.3**

### Property 8: 搜索结果完整性

*For any* 搜索词，高亮的匹配数量应等于文档中实际出现的次数
**Validates: Requirements 9.2**

### Property 9: 正则替换正确性

*For any* 有效的正则表达式和替换模式，替换结果应符合正则表达式语义
**Validates: Requirements 9.4**

### Property 10: 标签页状态一致性

*For any* 打开的文件，标签页显示的修改状态应与实际内容是否修改一致
**Validates: Requirements 8.4**

### Property 11: 自动保存完整性

*For any* 自动保存的内容，恢复后应与保存时的内容完全一致
**Validates: Requirements 10.4**

### Property 12: 文档统计准确性

*For any* 文档内容，状态栏显示的字数、字符数、行数应与实际内容一致
**Validates: Requirements 12.1, 12.2, 12.3**

## Error Handling

### 错误分类和处理策略

| 错误类型 | 处理策略 | 用户提示 |
| -------- | -------- | -------- |
| 文件读取失败 | 显示错误对话框，提供重试选项 | "无法打开文件: {原因}" |
| 文件保存失败 | 自动备份到临时目录，提示用户 | "保存失败，已备份到: {路径}" |
| 导出失败 | 记录详细日志，显示错误位置 | "导出失败（第{行}行）: {原因}" |
| 渲染超时 | 显示加载指示器，允许取消 | "正在渲染大文档..." |
| 快捷键冲突 | 高亮冲突项，提供替代建议 | "快捷键已被{功能}使用" |
| 主题格式错误 | 使用默认主题，显示警告 | "主题格式无效，已使用默认主题" |

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
        auto_save_file = self.backup_dir / '.autosave.md'
        if auto_save_file.exists():
            return auto_save_file.read_text(encoding='utf-8')
        return None
    
    def get_recovery_info(self) -> dict:
        """获取恢复信息"""
        auto_save_file = self.backup_dir / '.autosave.md'
        if auto_save_file.exists():
            stat = auto_save_file.stat()
            return {
                'exists': True,
                'time': datetime.fromtimestamp(stat.st_mtime),
                'size': stat.st_size,
            }
        return {'exists': False}
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
│   ├── test_toolbar.py
│   ├── test_smart_editing.py
│   ├── test_scroll_sync.py
│   ├── test_theme.py
│   └── test_export.py
├── property/
│   ├── test_bracket_completion.py
│   ├── test_indent_properties.py
│   ├── test_scroll_sync_properties.py
│   ├── test_theme_properties.py
│   ├── test_search_properties.py
│   └── test_statistics_properties.py
└── integration/
    ├── test_export_workflow.py
    └── test_edit_workflow.py
```

### 属性测试示例

```python
from hypothesis import given, strategies as st, settings

class TestBracketCompletionProperties:
    """括号自动补全属性测试"""
    
    BRACKETS = ['(', '[', '{', '"', "'", '`', '（', '【', '「']
    
    @given(st.sampled_from(BRACKETS))
    @settings(max_examples=100)
    def test_bracket_auto_completion(self, bracket):
        """
        **Feature: best-markdown-editor, Property 1: 括号自动补全一致性**
        **Validates: Requirements 2.2**
        
        对于任意括号字符，输入后应自动补全配对字符
        """
        editor = MockEditor()
        smart_edit = SmartEditingEnhanced(editor)
        
        initial_pos = editor.get_cursor_position()
        smart_edit.on_bracket_input(bracket)
        
        content = editor.get_content()
        final_pos = editor.get_cursor_position()
        
        # 验证配对字符已插入
        assert bracket in content
        assert SmartEditingEnhanced.BRACKET_PAIRS[bracket] in content
        
        # 验证光标在配对字符之间
        assert final_pos == initial_pos + 1


class TestSearchProperties:
    """搜索功能属性测试"""
    
    @given(
        st.text(min_size=1, max_size=100),
        st.text(min_size=1, max_size=20)
    )
    @settings(max_examples=100)
    def test_search_result_completeness(self, document, search_term):
        """
        **Feature: best-markdown-editor, Property 8: 搜索结果完整性**
        **Validates: Requirements 9.2**
        
        对于任意搜索词，高亮数量应等于实际出现次数
        """
        searcher = SearchEngine(document)
        results = searcher.find_all(search_term)
        
        # 计算实际出现次数
        actual_count = document.count(search_term)
        
        assert len(results) == actual_count


class TestStatisticsProperties:
    """文档统计属性测试"""
    
    @given(st.text(min_size=0, max_size=10000))
    @settings(max_examples=100)
    def test_statistics_accuracy(self, content):
        """
        **Feature: best-markdown-editor, Property 12: 文档统计准确性**
        **Validates: Requirements 12.1, 12.2, 12.3**
        
        对于任意文档内容，统计数据应准确
        """
        stats = EditorStatistics.calculate(content)
        
        # 验证行数
        expected_lines = content.count('\n') + 1 if content else 0
        assert stats.line_count == expected_lines
        
        # 验证字符数
        assert stats.char_count == len(content)
        
        # 验证字符数（不含空格）
        assert stats.char_count_no_spaces == len(content.replace(' ', '').replace('\n', '').replace('\t', ''))
```

