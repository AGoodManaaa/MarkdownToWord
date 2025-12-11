"""Windows DPI awareness utilities."""

import ctypes
import platform
from ..logging import log

def set_dpi_awareness():
    """
    设置进程的 DPI 感知模式。
    优先尝试 Windows 8.1+ 的 SetProcessDpiAwareness，
    降级使用 Windows Vista+ 的 SetProcessDPIAware。
    """
    try:
        # 尝试 Windows 8.1+ 的 SetProcessDpiAwareness
        # PROCESS_SYSTEM_DPI_AWARE = 1
        # PROCESS_PER_MONITOR_DPI_AWARE = 2
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        log("DPI awareness set to PROCESS_SYSTEM_DPI_AWARE via shcore")
    except (AttributeError, OSError):
        try:
            # 尝试 Windows Vista+ 的 SetProcessDPIAware
            ctypes.windll.user32.SetProcessDPIAware()
            log("DPI awareness set via user32.SetProcessDPIAware")
        except (AttributeError, OSError) as e:
            log(f"Failed to set DPI awareness: {e}")

def get_dpi_scale():
    """
    获取当前屏幕的 DPI 缩放比例。
    
    Returns:
        float: 缩放比例 (例如 1.0, 1.25, 1.5)
    """
    try:
        # 获取屏幕 DPI
        hdc = ctypes.windll.user32.GetDC(0)
        dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88) # 88 = LOGPIXELSX
        ctypes.windll.user32.ReleaseDC(0, hdc)
        return dpi / 96.0
    except Exception as e:
        log(f"Failed to get DPI scale: {e}")
        return 1.0