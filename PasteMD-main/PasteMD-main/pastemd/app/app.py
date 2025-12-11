"""Application entry point and initialization."""

import sys
import threading
import queue
import tkinter as tk

try:
    import ctypes
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("RichQAQ.PasteMD")
except Exception:
    pass

from ..utils.win32 import set_dpi_awareness
from .. import __version__
from ..core.state import app_state
from ..core.singleton import check_single_instance
from ..config.loader import ConfigLoader
from ..config.paths import get_app_icon_path
from ..utils.logging import log
from ..utils.version_checker import VersionChecker
from ..domains.notification.manager import NotificationManager
from ..i18n import DEFAULT_LANGUAGE, detect_system_language, set_language, t
from .wiring import Container


def initialize_application() -> Container:
    """初始化应用程序"""
    # 1. 加载配置
    config_loader = ConfigLoader()
    config = config_loader.load()
    app_state.config = config
    app_state.hotkey_str = config.get("hotkey", "<ctrl>+b")

    language_value = config.get("language", DEFAULT_LANGUAGE) or DEFAULT_LANGUAGE
    language = str(language_value)
    if language.lower() == DEFAULT_LANGUAGE:
        detected_language = detect_system_language()
        if detected_language and detected_language != DEFAULT_LANGUAGE:
            language = detected_language
            app_state.config["language"] = detected_language
            try:
                config_loader.save(app_state.config)
            except Exception as exc:
                log(f"Failed to persist auto-detected language: {exc}")
            log(f"Auto-detected system language: {detected_language}")
    set_language(language)
    
    # 2. 创建依赖注入容器
    container = Container()
    
    log("Application initialized successfully")
    return container


def show_startup_notification(notification_manager: NotificationManager) -> None:
    """显示启动通知"""
    try:
        # 确保图标路径存在（仅用于验证）
        get_app_icon_path()
        notification_manager.notify(
            "PasteMD",
            t("app.startup.success"),
            ok=True
        )
    except Exception as e:
        log(f"Failed to show startup notification: {e}")


def check_update_in_background(notification_manager: NotificationManager, tray_menu_manager=None) -> None:
    """在后台检查版本更新"""
    def _check():
        try:

            checker = VersionChecker(__version__)
            result = checker.check_update()
            
            if result and result.get("has_update"):
                latest_version = result.get("latest_version")
                release_url = result.get("release_url")
                
                # 使用菜单管理器的方法更新版本信息并重新绘制菜单
                if tray_menu_manager and app_state.icon:
                    tray_menu_manager.update_version_info(app_state.icon, latest_version, release_url)
                
                log(f"New version available: {latest_version}")
                log(f"Download URL: {release_url}")
        except Exception as e:
            log(f"Background version check failed: {e}")
    
    # 启动后台线程，不阻塞主程序
    thread = threading.Thread(target=_check, daemon=True)
    thread.start()


def main() -> None:
    """应用程序主入口点"""
    try:
        # 设置 DPI 感知（尽早调用）
        set_dpi_awareness()

        # 检查单实例运行
        if not check_single_instance():
            log("Application is already running")
            sys.exit(1)
        
        # 初始化应用程序
        container = initialize_application()

        # 初始化 UI 队列，确保 Tk 等 UI 操作始终在主线程
        ui_queue: queue.Queue = queue.Queue()
        app_state.ui_queue = ui_queue
        
        # 初始化退出事件
        import threading
        app_state.quit_event = threading.Event()
        
        # 初始化全局 Tk 实例 (解决 Tcl_AsyncDelete 问题)
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        app_state.root = root

        # 启动热键监听
        hotkey_runner = container.get_hotkey_runner()
        hotkey_runner.start()
        
        # 获取通知管理器和菜单管理器
        notification_manager = container.get_notification_manager()
        tray_menu_manager = container.tray_menu_manager
        
        # 显示启动通知
        show_startup_notification(notification_manager)
        
        # 启动后台版本检查（无需显示通知）
        check_update_in_background(notification_manager, tray_menu_manager)
        
        # 启动托盘（改为后台线程，避免阻塞主线程）
        tray_runner = container.get_tray_runner()
        threading.Thread(target=tray_runner.run, daemon=True).start()

        # UI 队列处理函数
        def process_ui_queue():
            try:
                # 检查退出事件
                quit_event = getattr(app_state, 'quit_event', None)
                if quit_event and quit_event.is_set():
                    # 退出事件被触发 - 清理所有窗口
                    _cleanup_all_windows()
                    root.quit()
                    return
                
                while True:
                    # 非阻塞获取任务
                    task = ui_queue.get_nowait()
                    if task is None:
                        # 退出信号 - 清理所有窗口
                        _cleanup_all_windows()
                        root.quit()
                        return
                    try:
                        task()
                    except Exception as e:
                        log(f"UI task error: {e}")
            except queue.Empty:
                pass
            finally:
                # 继续轮询 (100ms)
                root.after(100, process_ui_queue)
        
        def _cleanup_all_windows():
            """清理所有窗口"""
            try:
                # 销毁所有子窗口（Toplevel窗口）
                for widget in root.winfo_children():
                    try:
                        if hasattr(widget, 'destroy'):
                            widget.destroy()
                    except Exception as e:
                        log(f"Failed to destroy widget: {e}")
            except Exception as e:
                log(f"Error during window cleanup: {e}")

        # 启动队列处理
        root.after(100, process_ui_queue)
        
        # 进入主事件循环
        root.mainloop()
        
    except KeyboardInterrupt:
        log("Application interrupted by user")
    except Exception as e:
        log(f"Fatal error: {e}")
        raise
    finally:
        # 释放锁
        if app_state.instance_checker:
            app_state.instance_checker.release_lock()
        log("Application shutting down")


if __name__ == "__main__":
    main()
