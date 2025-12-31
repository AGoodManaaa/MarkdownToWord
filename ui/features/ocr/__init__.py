# -*- coding: utf-8 -*-
"""OCR 图片转 Markdown 功能模块"""

from .image_input import ImageInputManager, ImageLoadError
from .ocr_engine import OCREngine, OCRResult, OCRRegion, ContentType, RecognitionError, EngineInitError
from .markdown_gen import MarkdownGenerator
from .batch_processor import BatchOCRProcessor, BatchProgress
from .dialog import OCRDialog, OCRFeature

__all__ = [
    'ImageInputManager',
    'ImageLoadError',
    'OCREngine',
    'OCRResult',
    'OCRRegion',
    'ContentType',
    'RecognitionError',
    'EngineInitError',
    'MarkdownGenerator',
    'BatchOCRProcessor',
    'BatchProgress',
    'OCRDialog',
    'OCRFeature',
]
