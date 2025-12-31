# -*- coding: utf-8 -*-
"""图片输入管理模块 - 支持文件、剪贴板、截图"""

from typing import Optional, List, Callable, Set
from pathlib import Path
import io
import os

try:
    from PIL import Image, ImageGrab
except ImportError:
    Image = None
    ImageGrab = None


class ImageLoadError(Exception):
    """图片加载错误"""
    pass


class ImageInputManager:
    """图片输入管理器，支持多种输入方式"""
    
    SUPPORTED_FORMATS: Set[str] = {'.png', '.jpg', '.jpeg', '.bmp', '.webp', '.gif'}
    MAX_IMAGE_SIZE = 50 * 1024 * 1024  # 50MB
    
    def __init__(self, app=None):
        self.app = app
        self._screenshot_callback: Optional[Callable] = None
    
    def load_from_file(self, file_path: str) -> Optional['Image.Image']:
        """从文件加载图片
        
        Args:
            file_path: 图片文件路径
            
        Returns:
            PIL Image 对象，加载失败返回 None
            
        Raises:
            ImageLoadError: 文件不存在、格式不支持或文件损坏
        """
        if Image is None:
            raise ImageLoadError("Pillow 库未安装")
        
        path = Path(file_path)
        
        # 检查文件存在
        if not path.exists():
            raise ImageLoadError(f"文件不存在: {file_path}")
        
        # 检查文件扩展名
        ext = path.suffix.lower()
        if ext not in self.SUPPORTED_FORMATS:
            raise ImageLoadError(f"不支持的图片格式: {ext}，支持: {', '.join(self.SUPPORTED_FORMATS)}")
        
        # 检查文件大小
        file_size = path.stat().st_size
        if file_size > self.MAX_IMAGE_SIZE:
            raise ImageLoadError(f"文件过大: {file_size / 1024 / 1024:.1f}MB，最大支持 50MB")
        
        try:
            image = Image.open(file_path)
            # 验证图片可读
            image.verify()
            # 重新打开（verify 后需要重新打开）
            image = Image.open(file_path)
            return image
        except Exception as e:
            raise ImageLoadError(f"无法加载图片: {e}")
    
    def load_from_clipboard(self) -> Optional['Image.Image']:
        """从剪贴板加载图片
        
        Returns:
            PIL Image 对象，剪贴板无图片返回 None
            
        Raises:
            ImageLoadError: 剪贴板访问失败
        """
        if ImageGrab is None:
            raise ImageLoadError("Pillow ImageGrab 不可用")
        
        try:
            image = ImageGrab.grabclipboard()
            if image is None:
                return None
            if isinstance(image, Image.Image):
                return image
            # 可能是文件路径列表
            if isinstance(image, list) and len(image) > 0:
                return self.load_from_file(image[0])
            return None
        except Exception as e:
            raise ImageLoadError(f"无法从剪贴板获取图片: {e}")
    
    def load_from_bytes(self, data: bytes) -> Optional['Image.Image']:
        """从字节数据加载图片
        
        Args:
            data: 图片字节数据
            
        Returns:
            PIL Image 对象
        """
        if Image is None:
            raise ImageLoadError("Pillow 库未安装")
        
        try:
            return Image.open(io.BytesIO(data))
        except Exception as e:
            raise ImageLoadError(f"无法解析图片数据: {e}")
    
    def load_multiple_files(self, file_paths: List[str]) -> List['Image.Image']:
        """批量加载图片
        
        Args:
            file_paths: 图片文件路径列表
            
        Returns:
            成功加载的 PIL Image 对象列表
        """
        images = []
        for path in file_paths:
            try:
                img = self.load_from_file(path)
                if img:
                    images.append(img)
            except ImageLoadError:
                continue  # 跳过加载失败的图片
        return images
    
    def validate_image(self, image: 'Image.Image') -> bool:
        """验证图片有效性
        
        Args:
            image: PIL Image 对象
            
        Returns:
            图片是否有效
        """
        if image is None:
            return False
        
        try:
            # 检查基本属性
            if image.width <= 0 or image.height <= 0:
                return False
            
            # 检查图片模式
            valid_modes = {'RGB', 'RGBA', 'L', 'P', '1', 'CMYK'}
            if image.mode not in valid_modes:
                return False
            
            # 尝试访问像素数据
            image.load()
            return True
        except Exception:
            return False
    
    def get_image_info(self, image: 'Image.Image') -> dict:
        """获取图片信息
        
        Args:
            image: PIL Image 对象
            
        Returns:
            包含图片信息的字典
        """
        if image is None:
            return {}
        
        return {
            'width': image.width,
            'height': image.height,
            'mode': image.mode,
            'format': image.format,
            'size_bytes': len(image.tobytes()) if image.mode in ('RGB', 'RGBA', 'L') else 0,
        }
    
    def convert_to_rgb(self, image: 'Image.Image') -> 'Image.Image':
        """将图片转换为 RGB 模式（OCR 需要）
        
        Args:
            image: PIL Image 对象
            
        Returns:
            RGB 模式的图片
        """
        if image is None:
            return None
        
        if image.mode == 'RGB':
            return image
        
        if image.mode == 'RGBA':
            # 创建白色背景
            background = Image.new('RGB', image.size, (255, 255, 255))
            background.paste(image, mask=image.split()[3])
            return background
        
        return image.convert('RGB')
    
    def resize_for_ocr(self, image: 'Image.Image', max_size: int = 4096) -> 'Image.Image':
        """调整图片大小以适合 OCR 处理
        
        Args:
            image: PIL Image 对象
            max_size: 最大边长
            
        Returns:
            调整后的图片
        """
        if image is None:
            return None
        
        width, height = image.size
        if width <= max_size and height <= max_size:
            return image
        
        # 计算缩放比例
        scale = min(max_size / width, max_size / height)
        new_width = int(width * scale)
        new_height = int(height * scale)
        
        return image.resize((new_width, new_height), Image.Resampling.LANCZOS)
    
    def start_screenshot(self, callback: Callable[['Image.Image'], None]) -> None:
        """启动截图模式
        
        Args:
            callback: 截图完成后的回调函数
        """
        self._screenshot_callback = callback
        # 截图功能需要在 GUI 层实现
        if self.app and hasattr(self.app, '_start_screenshot_mode'):
            self.app._start_screenshot_mode(self._on_screenshot_complete)
    
    def _on_screenshot_complete(self, image: 'Image.Image') -> None:
        """截图完成回调"""
        if self._screenshot_callback:
            self._screenshot_callback(image)
            self._screenshot_callback = None
