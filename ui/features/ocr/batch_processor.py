# -*- coding: utf-8 -*-
"""批量 OCR 处理模块"""

from dataclasses import dataclass, field
from typing import List, Callable, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import time

from .ocr_engine import OCREngine, OCRResult, RecognitionError
from .image_input import ImageInputManager, ImageLoadError


@dataclass
class BatchProgress:
    """批量处理进度"""
    total: int
    completed: int
    current_file: str
    errors: List[str] = field(default_factory=list)
    
    @property
    def percentage(self) -> float:
        """完成百分比"""
        if self.total == 0:
            return 100.0
        return (self.completed / self.total) * 100
    
    @property
    def is_complete(self) -> bool:
        """是否完成"""
        return self.completed >= self.total


@dataclass
class BatchResult:
    """批量处理结果"""
    results: List[OCRResult]
    errors: List[str]
    total_time: float
    
    @property
    def success_count(self) -> int:
        """成功数量"""
        return len(self.results)
    
    @property
    def error_count(self) -> int:
        """失败数量"""
        return len(self.errors)


class BatchOCRProcessor:
    """批量 OCR 处理器"""
    
    def __init__(self, ocr_engine: Optional[OCREngine] = None, max_workers: int = 4):
        """初始化批量处理器
        
        Args:
            ocr_engine: OCR 引擎实例，如果为 None 则创建新实例
            max_workers: 最大并行工作线程数
        """
        self.ocr_engine = ocr_engine or OCREngine()
        self.max_workers = max_workers
        self.image_input = ImageInputManager()
        self._executor: Optional[ThreadPoolExecutor] = None
        self._cancelled = False
        self._lock = threading.Lock()
    
    def process_batch(
        self,
        image_paths: List[str],
        on_progress: Optional[Callable[[BatchProgress], None]] = None,
        on_complete: Optional[Callable[[BatchResult], None]] = None,
        detect_tables: bool = True,
        detect_formulas: bool = True
    ) -> BatchResult:
        """批量处理图片
        
        Args:
            image_paths: 图片路径列表
            on_progress: 进度回调函数
            on_complete: 完成回调函数
            detect_tables: 是否检测表格
            detect_formulas: 是否检测公式
            
        Returns:
            BatchResult 批量处理结果
        """
        self._cancelled = False
        start_time = time.time()
        
        results: List[OCRResult] = []
        errors: List[str] = []
        
        # 确保引擎已初始化
        if not self.ocr_engine.is_initialized:
            try:
                self.ocr_engine.initialize()
            except Exception as e:
                error_msg = f"OCR 引擎初始化失败: {e}"
                errors.append(error_msg)
                result = BatchResult(results=[], errors=errors, total_time=0)
                if on_complete:
                    on_complete(result)
                return result
        
        total = len(image_paths)
        completed = 0
        
        # 按顺序处理以保持顺序
        for i, path in enumerate(image_paths):
            if self._cancelled:
                errors.append("处理已取消")
                break
            
            # 更新进度
            progress = BatchProgress(
                total=total,
                completed=completed,
                current_file=path,
                errors=errors.copy()
            )
            if on_progress:
                on_progress(progress)
            
            # 处理单张图片
            try:
                result = self._process_single(
                    path, 
                    detect_tables=detect_tables,
                    detect_formulas=detect_formulas
                )
                result.source_image = path
                results.append(result)
            except Exception as e:
                errors.append(f"{path}: {e}")
            
            completed += 1
        
        # 最终进度
        final_progress = BatchProgress(
            total=total,
            completed=completed,
            current_file="",
            errors=errors.copy()
        )
        if on_progress:
            on_progress(final_progress)
        
        total_time = time.time() - start_time
        batch_result = BatchResult(
            results=results,
            errors=errors,
            total_time=total_time
        )
        
        if on_complete:
            on_complete(batch_result)
        
        return batch_result
    
    def process_batch_async(
        self,
        image_paths: List[str],
        on_progress: Optional[Callable[[BatchProgress], None]] = None,
        on_complete: Optional[Callable[[BatchResult], None]] = None,
        detect_tables: bool = True,
        detect_formulas: bool = True
    ) -> threading.Thread:
        """异步批量处理图片
        
        Args:
            image_paths: 图片路径列表
            on_progress: 进度回调函数
            on_complete: 完成回调函数
            detect_tables: 是否检测表格
            detect_formulas: 是否检测公式
            
        Returns:
            处理线程
        """
        thread = threading.Thread(
            target=self.process_batch,
            args=(image_paths, on_progress, on_complete, detect_tables, detect_formulas),
            daemon=True
        )
        thread.start()
        return thread
    
    def cancel(self) -> None:
        """取消批量处理"""
        with self._lock:
            self._cancelled = True
    
    def is_cancelled(self) -> bool:
        """是否已取消"""
        with self._lock:
            return self._cancelled
    
    def _process_single(
        self, 
        image_path: str,
        detect_tables: bool = True,
        detect_formulas: bool = True
    ) -> OCRResult:
        """处理单张图片
        
        Args:
            image_path: 图片路径
            detect_tables: 是否检测表格
            detect_formulas: 是否检测公式
            
        Returns:
            OCRResult 识别结果
            
        Raises:
            ImageLoadError: 图片加载失败
            RecognitionError: 识别失败
        """
        # 加载图片
        image = self.image_input.load_from_file(image_path)
        if image is None:
            raise ImageLoadError(f"无法加载图片: {image_path}")
        
        # 验证图片
        if not self.image_input.validate_image(image):
            raise ImageLoadError(f"图片无效: {image_path}")
        
        # 预处理
        image = self.image_input.convert_to_rgb(image)
        image = self.image_input.resize_for_ocr(image)
        
        # OCR 识别
        result = self.ocr_engine.recognize(
            image,
            detect_tables=detect_tables,
            detect_formulas=detect_formulas
        )
        
        return result
    
    def process_parallel(
        self,
        image_paths: List[str],
        on_progress: Optional[Callable[[BatchProgress], None]] = None,
        on_complete: Optional[Callable[[BatchResult], None]] = None
    ) -> BatchResult:
        """并行批量处理（不保证顺序）
        
        注意：此方法使用多线程并行处理，结果顺序可能与输入顺序不同。
        如果需要保持顺序，请使用 process_batch 方法。
        
        Args:
            image_paths: 图片路径列表
            on_progress: 进度回调函数
            on_complete: 完成回调函数
            
        Returns:
            BatchResult 批量处理结果
        """
        self._cancelled = False
        start_time = time.time()
        
        results: List[tuple] = []  # (index, result)
        errors: List[str] = []
        
        # 确保引擎已初始化
        if not self.ocr_engine.is_initialized:
            try:
                self.ocr_engine.initialize()
            except Exception as e:
                errors.append(f"OCR 引擎初始化失败: {e}")
                result = BatchResult(results=[], errors=errors, total_time=0)
                if on_complete:
                    on_complete(result)
                return result
        
        total = len(image_paths)
        completed = 0
        
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            self._executor = executor
            
            # 提交所有任务
            future_to_index = {
                executor.submit(self._process_single, path): (i, path)
                for i, path in enumerate(image_paths)
            }
            
            # 收集结果
            for future in as_completed(future_to_index):
                if self._cancelled:
                    break
                
                index, path = future_to_index[future]
                
                try:
                    result = future.result()
                    result.source_image = path
                    results.append((index, result))
                except Exception as e:
                    errors.append(f"{path}: {e}")
                
                completed += 1
                
                # 更新进度
                progress = BatchProgress(
                    total=total,
                    completed=completed,
                    current_file=path,
                    errors=errors.copy()
                )
                if on_progress:
                    on_progress(progress)
        
        self._executor = None
        
        # 按原始顺序排序结果
        results.sort(key=lambda x: x[0])
        ordered_results = [r[1] for r in results]
        
        total_time = time.time() - start_time
        batch_result = BatchResult(
            results=ordered_results,
            errors=errors,
            total_time=total_time
        )
        
        if on_complete:
            on_complete(batch_result)
        
        return batch_result
