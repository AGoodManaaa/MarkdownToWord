# Design Document: OCR 图片转 Markdown

## Overview

本设计文档描述了 OCR 图片转 Markdown 功能的技术实现方案。该功能使用 PaddleOCR 进行文字识别，支持表格结构识别和数学公式转 LaTeX。

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                      OCR Feature                             │
├─────────────────────────────────────────────────────────────┤
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐         │
│  │ ImageInput  │  │ OCREngine   │  │ ResultEditor│         │
│  │   Module    │──│   Module    │──│   Module    │         │
│  └─────────────┘  └─────────────┘  └─────────────┘         │
│         │               │                │                  │
│         ▼               ▼                ▼                  │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐         │
│  │ Screenshot  │  │ TextOCR     │  │ Markdown    │         │
│  │ Clipboard   │  │ TableOCR    │  │ Generator   │         │
│  │ FileLoader  │  │ FormulaOCR  │  │ Preview     │         │
│  └─────────────┘  └─────────────┘  └─────────────┘         │
└─────────────────────────────────────────────────────────────┘
```

## Components and Interfaces

### 1. ImageInputManager - 图片输入管理

```python
from typing import Optional, List, Callable
from PIL import Image
import io

class ImageInputManager:
    """图片输入管理器，支持多种输入方式"""
    
    SUPPORTED_FORMATS = {'.png', '.jpg', '.jpeg', '.bmp', '.webp', '.gif'}
    
    def __init__(self, app):
        self.app = app
        self._screenshot_callback: Optional[Callable] = None
    
    def load_from_file(self, file_path: str) -> Optional[Image.Image]:
        """从文件加载图片"""
        pass
    
    def load_from_clipboard(self) -> Optional[Image.Image]:
        """从剪贴板加载图片"""
        pass
    
    def start_screenshot(self, callback: Callable[[Image.Image], None]) -> None:
        """启动截图模式"""
        pass
    
    def load_multiple_files(self, file_paths: List[str]) -> List[Image.Image]:
        """批量加载图片"""
        pass
    
    def validate_image(self, image: Image.Image) -> bool:
        """验证图片有效性"""
        pass
```

### 2. OCREngine - OCR 识别引擎

```python
from dataclasses import dataclass
from typing import List, Tuple, Optional
from enum import Enum

class ContentType(Enum):
    TEXT = "text"
    TABLE = "table"
    FORMULA = "formula"
    IMAGE = "image"

@dataclass
class OCRRegion:
    """OCR 识别区域"""
    bbox: Tuple[int, int, int, int]  # x1, y1, x2, y2
    content_type: ContentType
    content: str
    confidence: float
    raw_data: Optional[dict] = None

@dataclass
class OCRResult:
    """OCR 识别结果"""
    regions: List[OCRRegion]
    markdown: str
    source_image: Optional[str] = None
    processing_time: float = 0.0

class OCREngine:
    """OCR 识别引擎"""
    
    def __init__(self, use_gpu: bool = False):
        self.use_gpu = use_gpu
        self._text_ocr = None
        self._table_ocr = None
        self._formula_ocr = None
        self._initialized = False
    
    def initialize(self) -> None:
        """初始化 OCR 引擎（延迟加载）"""
        pass
    
    def recognize(self, image: 'Image.Image') -> OCRResult:
        """识别图片内容"""
        pass
    
    def recognize_text(self, image: 'Image.Image') -> List[OCRRegion]:
        """识别文字"""
        pass
    
    def recognize_table(self, image: 'Image.Image') -> List[OCRRegion]:
        """识别表格"""
        pass
    
    def recognize_formula(self, image: 'Image.Image') -> List[OCRRegion]:
        """识别公式"""
        pass
    
    def _detect_content_types(self, image: 'Image.Image') -> List[Tuple[ContentType, Tuple]]:
        """检测图片中的内容类型和区域"""
        pass
```

### 3. MarkdownGenerator - Markdown 生成器

```python
class MarkdownGenerator:
    """将 OCR 结果转换为 Markdown"""
    
    def __init__(self):
        self.indent = "  "
    
    def generate(self, result: OCRResult) -> str:
        """生成完整的 Markdown 文档"""
        pass
    
    def text_to_markdown(self, region: OCRRegion) -> str:
        """文字转 Markdown"""
        pass
    
    def table_to_markdown(self, region: OCRRegion) -> str:
        """表格转 Markdown"""
        pass
    
    def formula_to_markdown(self, region: OCRRegion, inline: bool = False) -> str:
        """公式转 Markdown LaTeX"""
        pass
    
    def _detect_list_structure(self, text: str) -> str:
        """检测并转换列表结构"""
        pass
    
    def _detect_heading_structure(self, text: str) -> str:
        """检测并转换标题结构"""
        pass
```

### 4. OCRDialog - OCR 对话框界面

```python
class OCRDialog:
    """OCR 功能对话框"""
    
    def __init__(self, app):
        self.app = app
        self.image_input = ImageInputManager(app)
        self.ocr_engine = OCREngine()
        self.markdown_gen = MarkdownGenerator()
        self.current_image = None
        self.current_result = None
    
    def show(self) -> None:
        """显示 OCR 对话框"""
        pass
    
    def _create_ui(self) -> None:
        """创建界面"""
        pass
    
    def _on_file_select(self) -> None:
        """文件选择回调"""
        pass
    
    def _on_clipboard_paste(self) -> None:
        """剪贴板粘贴回调"""
        pass
    
    def _on_screenshot(self) -> None:
        """截图回调"""
        pass
    
    def _on_recognize(self) -> None:
        """开始识别"""
        pass
    
    def _on_insert(self) -> None:
        """插入到文档"""
        pass
    
    def _on_copy(self) -> None:
        """复制到剪贴板"""
        pass
```

### 5. BatchOCRProcessor - 批量处理器

```python
from concurrent.futures import ThreadPoolExecutor
from typing import Callable

@dataclass
class BatchProgress:
    """批量处理进度"""
    total: int
    completed: int
    current_file: str
    errors: List[str]

class BatchOCRProcessor:
    """批量 OCR 处理器"""
    
    def __init__(self, ocr_engine: OCREngine, max_workers: int = 4):
        self.ocr_engine = ocr_engine
        self.max_workers = max_workers
        self._executor = None
        self._cancelled = False
    
    def process_batch(
        self,
        images: List[str],
        on_progress: Callable[[BatchProgress], None],
        on_complete: Callable[[List[OCRResult]], None]
    ) -> None:
        """批量处理图片"""
        pass
    
    def cancel(self) -> None:
        """取消批量处理"""
        pass
    
    def _process_single(self, image_path: str) -> OCRResult:
        """处理单张图片"""
        pass
```

## Data Models

### OCR 配置

```python
ocr_config = {
    'engine': {
        'use_gpu': False,
        'lang': 'ch',  # ch, en, ch_en
        'det_model': 'ch_PP-OCRv4_det',
        'rec_model': 'ch_PP-OCRv4_rec',
        'table_model': 'ch_ppstructure_mobile_v2.0_SLANet',
        'formula_model': 'latex_ocr'
    },
    'recognition': {
        'min_confidence': 0.7,
        'detect_tables': True,
        'detect_formulas': True,
        'preserve_layout': True
    },
    'output': {
        'include_confidence': False,
        'include_source_comment': True,
        'table_style': 'github'  # github, simple
    }
}
```

## Correctness Properties

*A property is a characteristic or behavior that should hold true across all valid executions of a system-essentially, a formal statement about what the system should do. Properties serve as the bridge between human-readable specifications and machine-verifiable correctness guarantees.*

### Property 1: 图片格式验证

*For any* 图片文件路径，如果文件扩展名在支持列表中且文件可读，则 load_from_file 应返回有效的 Image 对象
**Validates: Requirements 1.5**

### Property 2: OCR 结果完整性

*For any* 有效图片，OCR 识别后的结果应包含至少一个区域，且每个区域的 confidence 在 0-1 之间
**Validates: Requirements 2.1, 2.4**

### Property 3: Markdown 表格语法正确性

*For any* 识别出的表格数据，生成的 Markdown 表格应符合 GitHub Flavored Markdown 语法规范
**Validates: Requirements 3.3**

### Property 4: LaTeX 公式语法正确性

*For any* 识别出的数学公式，生成的 LaTeX 应能被标准 LaTeX 解析器解析
**Validates: Requirements 4.2, 4.3**

### Property 5: 批量处理顺序保持

*For any* 批量处理的图片列表，输出结果的顺序应与输入顺序一致
**Validates: Requirements 5.4**

### Property 6: 批量处理容错性

*For any* 批量处理中的单个失败，不应影响其他图片的处理，且失败信息应被记录
**Validates: Requirements 5.5**

## Error Handling

```python
class OCRError(Exception):
    """OCR 相关错误基类"""
    pass

class ImageLoadError(OCRError):
    """图片加载错误"""
    pass

class RecognitionError(OCRError):
    """识别错误"""
    pass

class EngineInitError(OCRError):
    """引擎初始化错误"""
    pass

def handle_ocr_error(error: OCRError) -> str:
    """处理 OCR 错误，返回用户友好的消息"""
    error_messages = {
        ImageLoadError: "无法加载图片，请检查文件格式和路径",
        RecognitionError: "识别失败，请尝试使用更清晰的图片",
        EngineInitError: "OCR 引擎初始化失败，请检查依赖是否安装"
    }
    return error_messages.get(type(error), str(error))
```

## Testing Strategy

### 依赖库

使用 `pytest` 和 `hypothesis` 进行测试：

```python
# requirements-dev.txt
pytest>=7.0.0
hypothesis>=6.0.0
pillow>=9.0.0
```

### 测试覆盖

| 组件 | 单元测试 | 属性测试 |
| ---- | -------- | -------- |
| ImageInputManager | ✓ | ✓ (Property 1) |
| OCREngine | ✓ | ✓ (Property 2) |
| MarkdownGenerator | ✓ | ✓ (Property 3, 4) |
| BatchOCRProcessor | ✓ | ✓ (Property 5, 6) |

## Dependencies

```
paddlepaddle>=2.5.0  # 或 paddlepaddle-gpu
paddleocr>=2.7.0
pix2tex>=0.1.0  # LaTeX OCR
pillow>=9.0.0
```
