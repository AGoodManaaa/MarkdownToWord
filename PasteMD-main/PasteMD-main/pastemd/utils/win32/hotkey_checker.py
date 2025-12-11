"""Windows hotkey availability checker."""

import ctypes
from typing import Dict, Optional, Set, Tuple

from ...i18n import t
from ...utils.logging import log

user32 = ctypes.windll.user32

# Modifiers
MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008
MOD_NOREPEAT = 0x4000

MODIFIER_KEYS = {"ctrl", "shift", "alt", "cmd", "win"}

SYSTEM_HOTKEY_DESCRIPTIONS = {
    ('ctrl', 'c'): "hotkey.recorder.system.ctrl_c",
    ('ctrl', 'v'): "hotkey.recorder.system.ctrl_v",
    ('ctrl', 'x'): "hotkey.recorder.system.ctrl_x",
    ('ctrl', 'z'): "hotkey.recorder.system.ctrl_z",
    ('ctrl', 'y'): "hotkey.recorder.system.ctrl_y",
    ('ctrl', 'a'): "hotkey.recorder.system.ctrl_a",
    ('ctrl', 's'): "hotkey.recorder.system.ctrl_s",
    ('ctrl', 'f'): "hotkey.recorder.system.ctrl_f",
    ('ctrl', 'p'): "hotkey.recorder.system.ctrl_p",
    ('ctrl', 'n'): "hotkey.recorder.system.ctrl_n",
    ('ctrl', 'w'): "hotkey.recorder.system.ctrl_w",
    ('ctrl', 't'): "hotkey.recorder.system.ctrl_t",
    ('alt', 'f4'): "hotkey.recorder.system.alt_f4",
}

# Virtual Key Codes Mapping
VK_MAP: Dict[str, int] = {
    'backspace': 0x08, 'tab': 0x09, 'enter': 0x0D, 'pause': 0x13, 'caps_lock': 0x14, 'esc': 0x1B,
    'space': 0x20, 'page_up': 0x21, 'page_down': 0x22, 'end': 0x23, 'home': 0x24,
    'left': 0x25, 'up': 0x26, 'right': 0x27, 'down': 0x28,
    'print_screen': 0x2C, 'insert': 0x2D, 'delete': 0x2E,
    '0': 0x30, '1': 0x31, '2': 0x32, '3': 0x33, '4': 0x34,
    '5': 0x35, '6': 0x36, '7': 0x37, '8': 0x38, '9': 0x39,
    'a': 0x41, 'b': 0x42, 'c': 0x43, 'd': 0x44, 'e': 0x45,
    'f': 0x46, 'g': 0x47, 'h': 0x48, 'i': 0x49, 'j': 0x4A,
    'k': 0x4B, 'l': 0x4C, 'm': 0x4D, 'n': 0x4E, 'o': 0x4F,
    'p': 0x50, 'q': 0x51, 'r': 0x52, 's': 0x53, 't': 0x54,
    'u': 0x55, 'v': 0x56, 'w': 0x57, 'x': 0x58, 'y': 0x59, 'z': 0x5A,
    'f1': 0x70, 'f2': 0x71, 'f3': 0x72, 'f4': 0x73, 'f5': 0x74,
    'f6': 0x75, 'f7': 0x76, 'f8': 0x77, 'f9': 0x78, 'f10': 0x79,
    'f11': 0x7A, 'f12': 0x7B,
    'num0': 0x60, 'num1': 0x61, 'num2': 0x62, 'num3': 0x63, 'num4': 0x64,
    'num5': 0x65, 'num6': 0x66, 'num7': 0x67, 'num8': 0x68, 'num9': 0x69,
    'multiply': 0x6A, 'add': 0x6B, 'separator': 0x6C, 'subtract': 0x6D, 'decimal': 0x6E, 'divide': 0x6F,
    ';': 0xBA, '=': 0xBB, ',': 0xBC, '-': 0xBD, '.': 0xBE, '/': 0xBF, '`': 0xC0,
    '[': 0xDB, '\\': 0xDC, ']': 0xDD, "'": 0xDE,
}


class HotkeyChecker:
    """Windows Hotkey Availability Checker"""

    @staticmethod
    def _split_hotkey_keys(hotkey_str: str) -> Set[str]:
        return {
            part.strip()
            for part in hotkey_str.lower().replace("<", "").replace(">", "").split("+")
            if part.strip()
        }

    @staticmethod
    def validate_hotkey_keys(
        keys: Set[str],
        *,
        hotkey_repr: str = "",
        detailed: bool = False,
    ) -> Optional[str]:
        """
        Validate hotkey keys that have already been parsed into a set.
        Returns a localized error message if invalid; otherwise None.
        """
        has_modifier = bool(keys & MODIFIER_KEYS)
        has_normal_key = bool(keys - MODIFIER_KEYS)

        if not has_modifier:
            return t("hotkey.recorder.error.no_modifier")

        if not has_normal_key:
            return t("hotkey.recorder.error.no_normal_key")

        if keys & MODIFIER_KEYS == {"shift"}:
            translation_key = "hotkey.recorder.error.shift_only_long" if detailed else "hotkey.recorder.error.shift_only_short"
            return t(translation_key)

        for combo, desc_key in SYSTEM_HOTKEY_DESCRIPTIONS.items():
            if keys == set(combo):
                if detailed:
                    return t("hotkey.recorder.error.system_reserved_long", combo=t(desc_key))
                combo_text = hotkey_repr.upper() if hotkey_repr else "+".join(combo).upper()
                return t("hotkey.recorder.error.system_reserved_short", combo=combo_text)

        return None

    @staticmethod
    def validate_hotkey_string(hotkey_str: str, *, detailed: bool = False) -> Optional[str]:
        """
        Validate a pynput style hotkey string. Returns error text or None.
        """
        try:
            keys = HotkeyChecker._split_hotkey_keys(hotkey_str)
            return HotkeyChecker.validate_hotkey_keys(
                keys,
                hotkey_repr=hotkey_str,
                detailed=detailed,
            )
        except Exception as e:
            return t("hotkey.recorder.error.invalid_format", error=str(e))

    @staticmethod
    def parse_hotkey(hotkey_str: str) -> Optional[Tuple[int, int]]:
        """
        Parse pynput style hotkey string to (modifiers, vk_code).
        Example: "<ctrl>+<alt>+a" -> (MOD_CONTROL | MOD_ALT, 0x41)
        
        Args:
            hotkey_str: Hotkey string in pynput format
            
        Returns:
            Tuple of (modifiers, vk_code) or None if parsing fails
        """
        try:
            parts = hotkey_str.lower().replace("<", "").replace(">", "").split("+")
            modifiers = 0
            vk_code = 0
            
            for part in parts:
                part = part.strip()
                if not part:
                    continue
                    
                if part == 'ctrl':
                    modifiers |= MOD_CONTROL
                elif part == 'alt':
                    modifiers |= MOD_ALT
                elif part == 'shift':
                    modifiers |= MOD_SHIFT
                elif part in ('cmd', 'win'):
                    modifiers |= MOD_WIN
                else:
                    # It's a key
                    if part in VK_MAP:
                        vk_code = VK_MAP[part]
                    elif len(part) == 1:
                        # Try to get VK code for single character
                        res = user32.VkKeyScanW(ord(part))
                        if res != -1:
                            vk_code = res & 0xFF
                    
            if vk_code == 0:
                return None
                
            return modifiers, vk_code
        except Exception as e:
            log(f"Error parsing hotkey {hotkey_str}: {e}")
            return None

    @staticmethod
    def is_hotkey_available(hotkey_str: str) -> bool:
        """
        Check if a hotkey is available by trying to register it.
        
        Args:
            hotkey_str: Hotkey string in pynput format
            
        Returns:
            True if available (registration successful), False otherwise.
        """
        parsed = HotkeyChecker.parse_hotkey(hotkey_str)
        if not parsed:
            log(f"Could not parse hotkey for checking: {hotkey_str}")
            # If we can't parse it, we can't check it. 
            # Assume it's "available" in the sense that we won't block it due to conflict,
            # but validation elsewhere might catch it.
            return True 
            
        modifiers, vk = parsed
        
        # Use a random ID for testing (1 is fine for a quick check on NULL hwnd)
        hotkey_id = 1 
        
        # Try to register
        # hWnd = None (NULL) associates with current thread
        # If successful, it means the hotkey was not taken.
        success = user32.RegisterHotKey(None, hotkey_id, modifiers, vk)
        
        if success:
            # It was available, unregister it immediately
            user32.UnregisterHotKey(None, hotkey_id)
            return True
        else:
            # Failed to register, likely taken
            error_code = ctypes.get_last_error()
            log(f"Hotkey {hotkey_str} check failed. Error code: {error_code}")
            return False