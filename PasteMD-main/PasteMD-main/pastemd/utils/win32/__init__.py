"""Windows platform utilities."""

from .window import cleanup_background_wps_processes
from .hotkey_checker import HotkeyChecker
from .dpi import set_dpi_awareness, get_dpi_scale

__all__ = ['cleanup_background_wps_processes', 'HotkeyChecker', 'set_dpi_awareness', 'get_dpi_scale']
