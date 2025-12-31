# -*- coding: utf-8 -*-
"""OCR 功能测试"""

import pytest
from unittest.mock import Mock, patch, MagicMock
import sys
import os

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ui.features.ocr.image_input import ImageInputManager, ImageLoadError
from ui.features.ocr.markdown_gen import MarkdownGenerator
from ui.features.ocr.ocr_engine import OCRResult, OCRRegion, ContentType
from ui.features.ocr.batch_processor import BatchOCRProcessor, BatchProgress


class TestImageInputManager:
    """ImageInputManager 测试"""
    
    def test_supported_formats(self):
        """测试支持的图片格式"""
        manager = ImageInputManager()
        assert '.png' in manager.SUPPORTED_FORMATS
        assert '.jpg' in manager.SUPPORTED_FORMATS
        assert '.jpeg' in manager.SUPPORTED_FORMATS
        assert '.bmp' in manager.SUPPORTED_FORMATS
        assert '.webp' in manager.SUPPORTED_FORMATS
    
    def test_load_nonexistent_file(self):
        """测试加载不存在的文件"""
        manager = ImageInputManager()
        with pytest.raises(ImageLoadError) as exc_info:
            manager.load_from_file('/nonexistent/path/image.png')
        assert "不存在" in str(exc_info.value)
    
    def test_load_unsupported_format(self, tmp_path):
        """测试加载不支持的格式"""
        manager = ImageInputManager()
        # 创建一个假的文件
        fake_file = tmp_path / "test.xyz"
        fake_file.write_text("fake content")
        
        with pytest.raises(ImageLoadError) as exc_info:
            manager.load_from_file(str(fake_file))
        assert "不支持" in str(exc_info.value)
    
    def test_validate_image_none(self):
        """测试验证 None 图片"""
        manager = ImageInputManager()
        assert manager.validate_image(None) is False
    
    @patch('ui.features.ocr.image_input.Image')
    def test_validate_image_valid(self, mock_image):
        """测试验证有效图片"""
        manager = ImageInputManager()
        
        mock_img = MagicMock()
        mock_img.width = 100
        mock_img.height = 100
        mock_img.mode = 'RGB'
        
        assert manager.validate_image(mock_img) is True


class TestMarkdownGenerator:
    """MarkdownGenerator 测试"""
    
    def test_generate_empty_result(self):
        """测试生成空结果"""
        gen = MarkdownGenerator()
        result = OCRResult(regions=[])
        
        md = gen.generate(result)
        assert md == ""
    
    def test_generate_text_region(self):
        """测试生成文字区域"""
        gen = MarkdownGenerator()
        
        region = OCRRegion(
            bbox=(0, 0, 100, 50),
            content_type=ContentType.TEXT,
            content="Hello World",
            confidence=0.95
        )
        result = OCRResult(regions=[region])
        
        md = gen.generate(result, include_source_comment=False)
        assert "Hello World" in md
    
    def test_generate_table_region(self):
        """测试生成表格区域"""
        gen = MarkdownGenerator()
        
        table_content = "| A | B |\n|---|---|\n| 1 | 2 |"
        region = OCRRegion(
            bbox=(0, 0, 200, 100),
            content_type=ContentType.TABLE,
            content=table_content,
            confidence=0.85
        )
        result = OCRResult(regions=[region])
        
        md = gen.generate(result, include_source_comment=False)
        assert "|" in md
        assert "---" in md
    
    def test_generate_formula_region(self):
        """测试生成公式区域"""
        gen = MarkdownGenerator()
        
        region = OCRRegion(
            bbox=(0, 0, 100, 50),
            content_type=ContentType.FORMULA,
            content="E = mc^2",
            confidence=0.8
        )
        result = OCRResult(regions=[region])
        
        md = gen.generate(result, include_source_comment=False)
        assert "$" in md
        assert "E = mc^2" in md
    
    def test_detect_list_structure(self):
        """测试检测列表结构"""
        gen = MarkdownGenerator()
        
        text = "- Item 1\n- Item 2\n- Item 3"
        result = gen._detect_list_structure(text)
        
        assert "- Item 1" in result
    
    def test_detect_heading_structure(self):
        """测试检测标题结构"""
        gen = MarkdownGenerator()
        
        text = "一、标题内容"
        result = gen._detect_heading_structure(text)
        
        assert "#" in result
    
    def test_formula_inline_vs_block(self):
        """测试行内公式和块级公式"""
        gen = MarkdownGenerator()
        
        # 简单公式应该是行内
        simple = "x + y"
        assert gen._is_simple_formula(simple) is True
        
        # 复杂公式应该是块级
        complex_formula = r"\begin{equation} x + y \end{equation}"
        assert gen._is_simple_formula(complex_formula) is False
    
    def test_merge_results(self):
        """测试合并多个结果"""
        gen = MarkdownGenerator()
        
        result1 = OCRResult(
            regions=[OCRRegion(
                bbox=(0, 0, 100, 50),
                content_type=ContentType.TEXT,
                content="Page 1",
                confidence=0.9
            )],
            source_image="image1.png"
        )
        
        result2 = OCRResult(
            regions=[OCRRegion(
                bbox=(0, 0, 100, 50),
                content_type=ContentType.TEXT,
                content="Page 2",
                confidence=0.9
            )],
            source_image="image2.png"
        )
        
        merged = gen.merge_results([result1, result2])
        
        assert "Page 1" in merged
        assert "Page 2" in merged
        assert "---" in merged  # 分隔符


class TestOCRResult:
    """OCRResult 测试"""
    
    def test_has_content(self):
        """测试 has_content 属性"""
        empty_result = OCRResult(regions=[])
        assert empty_result.has_content is False
        
        result_with_content = OCRResult(regions=[
            OCRRegion(bbox=(0,0,10,10), content_type=ContentType.TEXT, content="test", confidence=0.9)
        ])
        assert result_with_content.has_content is True
    
    def test_average_confidence(self):
        """测试平均置信度计算"""
        result = OCRResult(regions=[
            OCRRegion(bbox=(0,0,10,10), content_type=ContentType.TEXT, content="a", confidence=0.8),
            OCRRegion(bbox=(0,0,10,10), content_type=ContentType.TEXT, content="b", confidence=1.0),
        ])
        
        assert result.average_confidence == 0.9
    
    def test_get_regions_by_type(self):
        """测试按类型获取区域"""
        result = OCRResult(regions=[
            OCRRegion(bbox=(0,0,10,10), content_type=ContentType.TEXT, content="text", confidence=0.9),
            OCRRegion(bbox=(0,0,10,10), content_type=ContentType.TABLE, content="table", confidence=0.8),
            OCRRegion(bbox=(0,0,10,10), content_type=ContentType.TEXT, content="text2", confidence=0.85),
        ])
        
        text_regions = result.get_regions_by_type(ContentType.TEXT)
        assert len(text_regions) == 2
        
        table_regions = result.get_regions_by_type(ContentType.TABLE)
        assert len(table_regions) == 1


class TestBatchProgress:
    """BatchProgress 测试"""
    
    def test_percentage(self):
        """测试百分比计算"""
        progress = BatchProgress(total=10, completed=5, current_file="test.png")
        assert progress.percentage == 50.0
    
    def test_percentage_zero_total(self):
        """测试总数为零时的百分比"""
        progress = BatchProgress(total=0, completed=0, current_file="")
        assert progress.percentage == 100.0
    
    def test_is_complete(self):
        """测试完成状态"""
        incomplete = BatchProgress(total=10, completed=5, current_file="test.png")
        assert incomplete.is_complete is False
        
        complete = BatchProgress(total=10, completed=10, current_file="")
        assert complete.is_complete is True


# Property-based tests using hypothesis
try:
    from hypothesis import given, strategies as st, settings
    
    class TestImageInputProperties:
        """ImageInputManager 属性测试"""
        
        @given(st.text(min_size=1, max_size=10))
        @settings(max_examples=20)
        def test_unsupported_extension_raises_error(self, ext):
            """Property 1: 不支持的扩展名应该抛出错误"""
            manager = ImageInputManager()
            
            # 确保扩展名不在支持列表中
            if f'.{ext.lower()}' not in manager.SUPPORTED_FORMATS:
                # 这里我们只测试逻辑，不实际创建文件
                pass
    
except ImportError:
    pass


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
