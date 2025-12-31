# -*- coding: utf-8 -*-
"""启动优化模块 - 延迟加载非核心功能，加速启动"""

import time
import threading
from typing import Callable, Dict, List, Optional, Tuple
from dataclasses import dataclass, field


@dataclass
class LazyModule:
    """延迟加载模块"""
    name: str
    loader: Callable
    priority: int = 5  # 1-10, 1 最高优先级
    loaded: bool = False
    load_time: float = 0
    error: Optional[str] = None


class StartupOptimizer:
    """启动优化器 - 管理延迟加载"""
    
    def __init__(self, app):
        self.app = app
        self._modules: Dict[str, LazyModule] = {}
        self._load_queue: List[str] = []
        self._loading = False
        self._start_time = time.time()
        self._core_loaded_time = 0
        self._all_loaded_time = 0
        
        # 加载阶段
        self._stages = {
            'core': [],      # 核心功能 (立即加载)
            'ui': [],        # UI 组件 (100ms 后)
            'features': [],  # 功能模块 (500ms 后)
            'optional': [],  # 可选功能 (1s 后)
        }
        
        # 进度回调
        self._progress_callback: Optional[Callable[[str, float], None]] = None
    
    def register_module(self, name: str, loader: Callable, 
                       stage: str = 'features', priority: int = 5):
        """注册延迟加载模块
        
        Args:
            name: 模块名称
            loader: 加载函数
            stage: 加载阶段 (core/ui/features/optional)
            priority: 优先级 (1-10)
        """
        module = LazyModule(name=name, loader=loader, priority=priority)
        self._modules[name] = module
        
        if stage in self._stages:
            self._stages[stage].append(name)
    
    def set_progress_callback(self, callback: Callable[[str, float], None]):
        """设置进度回调"""
        self._progress_callback = callback
    
    def load_core(self):
        """加载核心模块 (同步)"""
        self._load_stage('core')
        self._core_loaded_time = time.time() - self._start_time
    
    def load_deferred(self):
        """加载延迟模块 (异步)"""
        if self._loading:
            return
        
        self._loading = True
        
        def load_async():
            # UI 组件 - 100ms 后
            self.app.after(100, lambda: self._load_stage('ui'))
            
            # 功能模块 - 500ms 后
            self.app.after(500, lambda: self._load_stage('features'))
            
            # 可选功能 - 1s 后
            self.app.after(1000, lambda: self._load_stage('optional'))
            
            # 完成
            self.app.after(1500, self._on_all_loaded)
        
        thread = threading.Thread(target=load_async, daemon=True)
        thread.start()
    
    def _load_stage(self, stage: str):
        """加载指定阶段的模块"""
        modules = self._stages.get(stage, [])
        
        # 按优先级排序
        modules.sort(key=lambda n: self._modules[n].priority)
        
        total = len(modules)
        for i, name in enumerate(modules):
            self._load_module(name)
            
            # 更新进度
            if self._progress_callback:
                progress = (i + 1) / total if total > 0 else 1
                self._progress_callback(f"加载 {name}...", progress)
    
    def _load_module(self, name: str):
        """加载单个模块"""
        module = self._modules.get(name)
        if not module or module.loaded:
            return
        
        start = time.time()
        try:
            module.loader()
            module.loaded = True
            module.load_time = time.time() - start
        except Exception as e:
            module.error = str(e)
            print(f"⚠️ 模块 {name} 加载失败: {e}")
    
    def _on_all_loaded(self):
        """所有模块加载完成"""
        self._all_loaded_time = time.time() - self._start_time
        self._loading = False
        
        if self._progress_callback:
            self._progress_callback("就绪", 1.0)
    
    def is_loaded(self, name: str) -> bool:
        """检查模块是否已加载"""
        module = self._modules.get(name)
        return module.loaded if module else False
    
    def ensure_loaded(self, name: str):
        """确保模块已加载 (同步)"""
        if not self.is_loaded(name):
            self._load_module(name)
    
    def get_stats(self) -> Dict:
        """获取启动统计"""
        loaded = sum(1 for m in self._modules.values() if m.loaded)
        failed = sum(1 for m in self._modules.values() if m.error)
        
        return {
            'core_load_time': self._core_loaded_time,
            'total_load_time': self._all_loaded_time,
            'modules_total': len(self._modules),
            'modules_loaded': loaded,
            'modules_failed': failed,
            'module_times': {
                name: m.load_time 
                for name, m in self._modules.items() 
                if m.loaded
            }
        }


class SplashScreen:
    """启动画面"""
    
    def __init__(self, app):
        self.app = app
        self.window = None
        self._progress_var = None
        self._status_var = None
    
    def show(self):
        """显示启动画面"""
        import tkinter as tk
        
        self.window = tk.Toplevel(self.app)
        self.window.overrideredirect(True)  # 无边框
        self.window.attributes('-topmost', True)
        
        # 居中显示
        width, height = 400, 200
        screen_w = self.window.winfo_screenwidth()
        screen_h = self.window.winfo_screenheight()
        x = (screen_w - width) // 2
        y = (screen_h - height) // 2
        self.window.geometry(f"{width}x{height}+{x}+{y}")
        
        # 背景
        self.window.configure(bg='#10B981')
        
        # 标题
        title = tk.Label(
            self.window,
            text="✨ Markdown → Word",
            font=('Segoe UI', 24, 'bold'),
            fg='white',
            bg='#10B981'
        )
        title.pack(pady=(40, 10))
        
        # 副标题
        subtitle = tk.Label(
            self.window,
            text="正在启动...",
            font=('Segoe UI', 12),
            fg='white',
            bg='#10B981'
        )
        subtitle.pack()
        
        # 进度条
        from tkinter import ttk
        style = ttk.Style()
        style.configure("Splash.Horizontal.TProgressbar", 
                       background='white', 
                       troughcolor='#059669')
        
        self._progress_var = tk.DoubleVar(value=0)
        progress = ttk.Progressbar(
            self.window,
            variable=self._progress_var,
            maximum=100,
            length=300,
            style="Splash.Horizontal.TProgressbar"
        )
        progress.pack(pady=20)
        
        # 状态文本
        self._status_var = tk.StringVar(value="初始化...")
        status = tk.Label(
            self.window,
            textvariable=self._status_var,
            font=('Segoe UI', 10),
            fg='white',
            bg='#10B981'
        )
        status.pack()
        
        self.window.update()
    
    def update_progress(self, status: str, progress: float):
        """更新进度"""
        if self.window and self.window.winfo_exists():
            self._status_var.set(status)
            self._progress_var.set(progress * 100)
            self.window.update()
    
    def hide(self):
        """隐藏启动画面"""
        if self.window and self.window.winfo_exists():
            self.window.destroy()
            self.window = None


class MemoryOptimizer:
    """内存优化器"""
    
    def __init__(self):
        self._gc_interval = 60  # 秒
        self._last_gc = time.time()
        self._enabled = True
    
    def enable(self):
        """启用内存优化"""
        self._enabled = True
    
    def disable(self):
        """禁用内存优化"""
        self._enabled = False
    
    def optimize(self):
        """执行内存优化"""
        if not self._enabled:
            return
        
        import gc
        gc.collect()
        self._last_gc = time.time()
    
    def schedule_gc(self, app):
        """定时垃圾回收"""
        def gc_task():
            if time.time() - self._last_gc > self._gc_interval:
                self.optimize()
            app.after(self._gc_interval * 1000, gc_task)
        
        app.after(self._gc_interval * 1000, gc_task)
    
    def get_memory_usage(self) -> Dict:
        """获取内存使用情况"""
        try:
            import psutil
            process = psutil.Process()
            mem = process.memory_info()
            return {
                'rss_mb': mem.rss / 1024 / 1024,
                'vms_mb': mem.vms / 1024 / 1024,
            }
        except ImportError:
            return {'error': 'psutil not installed'}
        except Exception as e:
            return {'error': str(e)}


class CacheManager:
    """缓存管理器"""
    
    def __init__(self, max_size_mb: int = 100):
        self._caches: Dict[str, Dict] = {}
        self._max_size = max_size_mb * 1024 * 1024
        self._current_size = 0
    
    def create_cache(self, name: str, max_items: int = 1000) -> Dict:
        """创建缓存"""
        cache = LRUCache(max_items)
        self._caches[name] = cache
        return cache
    
    def get_cache(self, name: str) -> Optional[Dict]:
        """获取缓存"""
        return self._caches.get(name)
    
    def clear_cache(self, name: str):
        """清除指定缓存"""
        if name in self._caches:
            self._caches[name].clear()
    
    def clear_all(self):
        """清除所有缓存"""
        for cache in self._caches.values():
            cache.clear()
        self._current_size = 0
    
    def get_stats(self) -> Dict:
        """获取缓存统计"""
        stats = {}
        for name, cache in self._caches.items():
            stats[name] = {
                'items': len(cache),
                'hits': getattr(cache, '_hits', 0),
                'misses': getattr(cache, '_misses', 0),
            }
        return stats


class LRUCache:
    """LRU 缓存"""
    
    def __init__(self, max_items: int = 1000):
        self._max_items = max_items
        self._cache: Dict = {}
        self._order: List = []
        self._hits = 0
        self._misses = 0
    
    def get(self, key, default=None):
        """获取缓存项"""
        if key in self._cache:
            self._hits += 1
            # 移到最后（最近使用）
            self._order.remove(key)
            self._order.append(key)
            return self._cache[key]
        
        self._misses += 1
        return default
    
    def set(self, key, value):
        """设置缓存项"""
        if key in self._cache:
            self._order.remove(key)
        elif len(self._cache) >= self._max_items:
            # 移除最旧的项
            oldest = self._order.pop(0)
            del self._cache[oldest]
        
        self._cache[key] = value
        self._order.append(key)
    
    def __contains__(self, key):
        return key in self._cache
    
    def __len__(self):
        return len(self._cache)
    
    def clear(self):
        """清除缓存"""
        self._cache.clear()
        self._order.clear()
        self._hits = 0
        self._misses = 0
