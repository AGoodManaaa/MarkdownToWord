# -*- coding: utf-8 -*-
"""OCR 识别引擎模块 - 支持文字、表格、公式识别"""

from dataclasses import dataclass, field
from typing import List, Tuple, Optional, Dict, Any
from enum import Enum
import time

try:
    from PIL import Image
except ImportError:
    Image = None


class ContentType(Enum):
    """内容类型枚举"""
    TEXT = "text"
    TABLE = "table"
    FORMULA = "formula"
    IMAGE = "image"


class RecognitionError(Exception):
    """识别错误"""
    pass


class EngineInitError(Exception):
    """引擎初始化错误"""
    pass


@dataclass
class OCRRegion:
    """OCR 识别区域"""
    bbox: Tuple[int, int, int, int]  # x1, y1, x2, y2
    content_type: ContentType
    content: str
    confidence: float
    raw_data: Optional[Dict[str, Any]] = None
    
    def __post_init__(self):
        # 确保 confidence 在 0-1 之间
        self.confidence = max(0.0, min(1.0, self.confidence))


@dataclass
class OCRResult:
    """OCR 识别结果"""
    regions: List[OCRRegion] = field(default_factory=list)
    markdown: str = ""
    source_image: Optional[str] = None
    processing_time: float = 0.0
    
    @property
    def has_content(self) -> bool:
        """是否有识别内容"""
        return len(self.regions) > 0
    
    @property
    def average_confidence(self) -> float:
        """平均置信度"""
        if not self.regions:
            return 0.0
        return sum(r.confidence for r in self.regions) / len(self.regions)
    
    def get_regions_by_type(self, content_type: ContentType) -> List[OCRRegion]:
        """按类型获取区域"""
        return [r for r in self.regions if r.content_type == content_type]


class OCREngine:
    """OCR 识别引擎
    
    支持:
    - 文字识别 (PaddleOCR)
    - 表格识别 (PaddleOCR PP-Structure)
    - 公式识别 (pix2tex)
    """
    
    def __init__(self, use_gpu: bool = False, lang: str = 'ch'):
        """初始化 OCR 引擎
        
        Args:
            use_gpu: 是否使用 GPU 加速
            lang: 识别语言 ('ch', 'en', 'ch_en')
        """
        self.use_gpu = use_gpu
        self.lang = lang
        self._text_ocr = None
        self._table_ocr = None
        self._formula_ocr = None
        self._initialized = False
        self._init_error: Optional[str] = None
    
    def initialize(self) -> None:
        """初始化 OCR 引擎（延迟加载）
        
        Raises:
            EngineInitError: 初始化失败
        """
        if self._initialized:
            return
        
        errors = []
        
        # 初始化文字 OCR
        try:
            from paddleocr import PaddleOCR
            self._text_ocr = PaddleOCR(
                use_angle_cls=True,
                lang=self.lang,
                use_gpu=self.use_gpu,
                show_log=False
            )
        except ImportError:
            errors.append("PaddleOCR 未安装，请运行: pip install paddleocr")
        except Exception as e:
            errors.append(f"PaddleOCR 初始化失败: {e}")
        
        # 初始化表格 OCR
        try:
            from paddleocr import PPStructure
            self._table_ocr = PPStructure(
                table=True,
                ocr=True,
                show_log=False,
                use_gpu=self.use_gpu
            )
        except ImportError:
            pass  # 表格识别可选
        except Exception:
            pass
        
        # 初始化公式 OCR
        try:
            from pix2tex.cli import LatexOCR
            self._formula_ocr = LatexOCR()
        except ImportError:
            pass  # 公式识别可选
        except Exception:
            pass
        
        if self._text_ocr is None:
            self._init_error = "; ".join(errors) if errors else "OCR 引擎初始化失败"
            raise EngineInitError(self._init_error)
        
        self._initialized = True
    
    @property
    def is_initialized(self) -> bool:
        """是否已初始化"""
        return self._initialized
    
    @property
    def supports_table(self) -> bool:
        """是否支持表格识别"""
        return self._table_ocr is not None
    
    @property
    def supports_formula(self) -> bool:
        """是否支持公式识别"""
        return self._formula_ocr is not None
    
    def recognize(self, image: 'Image.Image', 
                  detect_tables: bool = True,
                  detect_formulas: bool = True) -> OCRResult:
        """识别图片内容
        
        Args:
            image: PIL Image 对象
            detect_tables: 是否检测表格
            detect_formulas: 是否检测公式
            
        Returns:
            OCRResult 识别结果
            
        Raises:
            RecognitionError: 识别失败
        """
        if not self._initialized:
            self.initialize()
        
        start_time = time.time()
        regions: List[OCRRegion] = []
        
        try:
            # 转换图片格式
            import numpy as np
            if image.mode != 'RGB':
                image = image.convert('RGB')
            img_array = np.array(image)
            
            # 文字识别
            text_regions = self.recognize_text(img_array)
            regions.extend(text_regions)
            
            # 表格识别
            if detect_tables and self._table_ocr:
                table_regions = self.recognize_table(img_array)
                regions.extend(table_regions)
            
            # 公式识别（如果检测到可能的公式区域）
            if detect_formulas and self._formula_ocr:
                formula_regions = self.recognize_formula(image)
                regions.extend(formula_regions)
            
            # 按位置排序（从上到下，从左到右）
            regions.sort(key=lambda r: (r.bbox[1], r.bbox[0]))
            
            processing_time = time.time() - start_time
            
            return OCRResult(
                regions=regions,
                processing_time=processing_time
            )
            
        except Exception as e:
            raise RecognitionError(f"识别失败: {e}")
    
    def recognize_text(self, img_array) -> List[OCRRegion]:
        """识别文字
        
        Args:
            img_array: numpy 数组格式的图片
            
        Returns:
            OCRRegion 列表
        """
        if self._text_ocr is None:
            return []
        
        regions = []
        try:
            result = self._text_ocr.ocr(img_array, cls=True)
            
            if result and result[0]:
                for line in result[0]:
                    if line and len(line) >= 2:
                        bbox_points = line[0]
                        text_info = line[1]
                        
                        # 转换 bbox 格式
                        x_coords = [p[0] for p in bbox_points]
                        y_coords = [p[1] for p in bbox_points]
                        bbox = (
                            int(min(x_coords)),
                            int(min(y_coords)),
                            int(max(x_coords)),
                            int(max(y_coords))
                        )
                        
                        text = text_info[0] if isinstance(text_info, tuple) else str(text_info)
                        confidence = text_info[1] if isinstance(text_info, tuple) and len(text_info) > 1 else 0.9
                        
                        regions.append(OCRRegion(
                            bbox=bbox,
                            content_type=ContentType.TEXT,
                            content=text,
                            confidence=float(confidence),
                            raw_data={'bbox_points': bbox_points}
                        ))
        except Exception:
            pass
        
        return regions
    
    def recognize_table(self, img_array) -> List[OCRRegion]:
        """识别表格
        
        Args:
            img_array: numpy 数组格式的图片
            
        Returns:
            OCRRegion 列表
        """
        if self._table_ocr is None:
            return []
        
        regions = []
        try:
            result = self._table_ocr(img_array)
            
            for item in result:
                if item.get('type') == 'table':
                    bbox = item.get('bbox', [0, 0, 0, 0])
                    html = item.get('res', {}).get('html', '')
                    
                    # 将 HTML 表格转换为 Markdown
                    md_table = self._html_table_to_markdown(html)
                    
                    if md_table:
                        regions.append(OCRRegion(
                            bbox=tuple(bbox),
                            content_type=ContentType.TABLE,
                            content=md_table,
                            confidence=0.85,
                            raw_data={'html': html}
                        ))
        except Exception:
            pass
        
        return regions
    
    def recognize_formula(self, image: 'Image.Image') -> List[OCRRegion]:
        """识别公式
        
        Args:
            image: PIL Image 对象
            
        Returns:
            OCRRegion 列表
        """
        if self._formula_ocr is None:
            return []
        
        regions = []
        try:
            # pix2tex 对整张图片进行公式识别
            latex = self._formula_ocr(image)
            
            if latex and latex.strip():
                regions.append(OCRRegion(
                    bbox=(0, 0, image.width, image.height),
                    content_type=ContentType.FORMULA,
                    content=latex.strip(),
                    confidence=0.8
                ))
        except Exception:
            pass
        
        return regions
    
    def _html_table_to_markdown(self, html: str) -> str:
        """将 HTML 表格转换为 Markdown 格式
        
        Args:
            html: HTML 表格字符串
            
        Returns:
            Markdown 表格字符串
        """
        if not html:
            return ""
        
        try:
            from lxml import etree
            
            # 解析 HTML
            parser = etree.HTMLParser()
            tree = etree.fromstring(html, parser)
            
            rows = []
            for tr in tree.xpath('//tr'):
                cells = []
                for td in tr.xpath('td|th'):
                    text = ''.join(td.itertext()).strip()
                    cells.append(text)
                if cells:
                    rows.append(cells)
            
            if not rows:
                return ""
            
            # 生成 Markdown 表格
            md_lines = []
            
            # 表头
            header = rows[0]
            md_lines.append('| ' + ' | '.join(header) + ' |')
            md_lines.append('| ' + ' | '.join(['---'] * len(header)) + ' |')
            
            # 数据行
            for row in rows[1:]:
                # 确保列数一致
                while len(row) < len(header):
                    row.append('')
                md_lines.append('| ' + ' | '.join(row[:len(header)]) + ' |')
            
            return '\n'.join(md_lines)
            
        except Exception:
            return ""
    
    def get_status(self) -> Dict[str, Any]:
        """获取引擎状态
        
        Returns:
            状态信息字典
        """
        return {
            'initialized': self._initialized,
            'use_gpu': self.use_gpu,
            'lang': self.lang,
            'supports_text': self._text_ocr is not None,
            'supports_table': self._table_ocr is not None,
            'supports_formula': self._formula_ocr is not None,
            'init_error': self._init_error
        }
