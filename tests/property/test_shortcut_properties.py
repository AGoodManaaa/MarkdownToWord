# -*- coding: utf-8 -*-
"""
快捷键唯一性属性测试
**Feature: markdown-editor-optimization, Property 9: 快捷键唯一性**
**Validates: Requirements 8.1, 8.3**
"""

import pytest
import re
from typing import Dict, List, Set, Tuple


def extract_shortcuts_from_file(filepath: str) -> List[Tuple[str, str, int]]:
    """
    从Python文件中提取快捷键绑定
    
    Returns:
        List of (shortcut, action/description, line_number)
    """
    shortcuts = []
    
    # 排除的事件类型（这些不是用户快捷键，可以多次绑定）
    excluded_events = {
        '<keyrelease>', '<keypress>', '<key>', '<button-1>', '<button-2>', '<button-3>',
        '<mousewheel>', '<configure>', '<focusin>', '<focusout>', '<enter>', '<leave>',
        '<<modified>>', '<<paste>>', '<<cut>>', '<<copy>>', '<motion>', '<buttonrelease>',
        '<escape>', '<tab>', '<return>', '<backspace>', '<delete>'
    }
    
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            lines = f.readlines()
    except FileNotFoundError:
        return shortcuts
    
    # 匹配 self.bind('<Control-...>', ...) 或 app.bind('<Control-...>', ...)
    bind_pattern = re.compile(
        r"\.bind\s*\(\s*['\"](<[^'\"]+>)['\"]"
    )
    
    for line_num, line in enumerate(lines, 1):
        match = bind_pattern.search(line)
        if match:
            shortcut = match.group(1)
            # 跳过非快捷键事件
            if shortcut.lower() in excluded_events:
                continue
            # 只关注包含 Control/Alt/Shift 修饰键的快捷键
            if not any(mod in shortcut.lower() for mod in ['control', 'alt', 'shift', 'f1', 'f2', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8', 'f9', 'f10', 'f11', 'f12']):
                continue
            # 提取注释作为描述
            comment_match = re.search(r'#\s*(.+)$', line)
            description = comment_match.group(1) if comment_match else line.strip()
            shortcuts.append((shortcut, description, line_num))
    
    return shortcuts


def normalize_shortcut(shortcut: str) -> str:
    """
    标准化快捷键字符串以便比较
    例如: '<Control-Shift-C>' 和 '<Control-Shift-c>' 应该被视为相同
    """
    # 转换为小写
    normalized = shortcut.lower()
    # 标准化修饰键顺序: control -> shift -> alt
    parts = normalized.strip('<>').split('-')
    modifiers = []
    key = parts[-1] if parts else ''
    
    for part in parts[:-1]:
        if part in ('control', 'ctrl'):
            modifiers.append('control')
        elif part == 'shift':
            modifiers.append('shift')
        elif part == 'alt':
            modifiers.append('alt')
        else:
            modifiers.append(part)
    
    # 排序修饰键
    modifiers.sort()
    return f"<{'-'.join(modifiers + [key])}>"


class TestShortcutProperties:
    """快捷键属性测试类"""
    
    def test_no_duplicate_shortcuts_in_gui(self):
        """
        **Feature: markdown-editor-optimization, Property 9: 快捷键唯一性**
        **Validates: Requirements 8.1, 8.3**
        
        测试gui.py中没有重复的快捷键绑定
        """
        shortcuts = extract_shortcuts_from_file('gui.py')
        
        # 按标准化后的快捷键分组
        shortcut_map: Dict[str, List[Tuple[str, str, int]]] = {}
        
        for shortcut, desc, line_num in shortcuts:
            normalized = normalize_shortcut(shortcut)
            if normalized not in shortcut_map:
                shortcut_map[normalized] = []
            shortcut_map[normalized].append((shortcut, desc, line_num))
        
        # 检查重复
        duplicates = {k: v for k, v in shortcut_map.items() if len(v) > 1}
        
        if duplicates:
            error_msg = "发现重复的快捷键绑定:\n"
            for shortcut, bindings in duplicates.items():
                error_msg += f"\n  {shortcut}:\n"
                for orig, desc, line in bindings:
                    error_msg += f"    - 第{line}行: {desc[:50]}...\n"
            pytest.fail(error_msg)
    
    def test_no_duplicate_shortcuts_in_keyboard_manager(self):
        """
        测试KeyboardShortcutsManager中没有重复的快捷键定义
        """
        from ui.keyboard_shortcuts import KeyboardShortcutsManager
        
        # 获取默认快捷键
        shortcuts = KeyboardShortcutsManager.DEFAULT_SHORTCUTS
        
        # 按标准化后的快捷键分组
        shortcut_map: Dict[str, List[str]] = {}
        
        for shortcut, action in shortcuts.items():
            normalized = normalize_shortcut(shortcut)
            if normalized not in shortcut_map:
                shortcut_map[normalized] = []
            shortcut_map[normalized].append(f"{shortcut} -> {action}")
        
        # 检查重复
        duplicates = {k: v for k, v in shortcut_map.items() if len(v) > 1}
        
        if duplicates:
            error_msg = "KeyboardShortcutsManager中发现重复的快捷键:\n"
            for shortcut, bindings in duplicates.items():
                error_msg += f"\n  {shortcut}:\n"
                for binding in bindings:
                    error_msg += f"    - {binding}\n"
            pytest.fail(error_msg)
    
    def test_ctrl_shift_c_not_duplicated(self):
        """
        特别测试: Ctrl+Shift+C 不应该被重复绑定
        这是之前发现的具体bug
        """
        shortcuts = extract_shortcuts_from_file('gui.py')
        
        ctrl_shift_c_bindings = [
            (s, d, l) for s, d, l in shortcuts 
            if normalize_shortcut(s) == '<control-shift-c>'
        ]
        
        assert len(ctrl_shift_c_bindings) <= 1, \
            f"Ctrl+Shift+C 被绑定了 {len(ctrl_shift_c_bindings)} 次: {ctrl_shift_c_bindings}"
    
    def test_collaboration_uses_ctrl_alt_c(self):
        """
        测试协作功能使用 Ctrl+Alt+C 而不是 Ctrl+Shift+C
        """
        shortcuts = extract_shortcuts_from_file('gui.py')
        
        # 查找协作相关的绑定
        collab_bindings = [
            (s, d, l) for s, d, l in shortcuts 
            if '协作' in d or 'collaboration' in d.lower()
        ]
        
        for shortcut, desc, line in collab_bindings:
            normalized = normalize_shortcut(shortcut)
            assert normalized != '<control-shift-c>', \
                f"协作功能仍然使用 Ctrl+Shift+C (第{line}行)"
            # 应该使用 Ctrl+Alt+C
            assert normalized == '<alt-c-control>' or 'alt' in normalized.lower(), \
                f"协作功能应该使用 Ctrl+Alt+C，但实际是 {shortcut} (第{line}行)"
    
    def test_all_shortcuts_have_unique_actions(self):
        """
        **Feature: markdown-editor-optimization, Property 9: 快捷键唯一性**
        **Validates: Requirements 8.1, 8.3**
        
        测试所有快捷键都映射到唯一的动作（没有同一个快捷键绑定到多个不同功能）
        """
        from ui.keyboard_shortcuts import KeyboardShortcutsManager
        
        shortcuts = KeyboardShortcutsManager.DEFAULT_SHORTCUTS
        
        # 检查每个快捷键只绑定到一个动作
        seen_shortcuts: Dict[str, str] = {}
        conflicts = []
        
        for shortcut, action in shortcuts.items():
            normalized = normalize_shortcut(shortcut)
            if normalized in seen_shortcuts:
                if seen_shortcuts[normalized] != action:
                    conflicts.append(
                        f"快捷键 {shortcut} ({normalized}) 绑定到多个动作: "
                        f"'{seen_shortcuts[normalized]}' 和 '{action}'"
                    )
            else:
                seen_shortcuts[normalized] = action
        
        assert not conflicts, "\n".join(conflicts)
    
    def test_shortcut_format_validity(self):
        """
        测试所有快捷键格式有效
        """
        from ui.keyboard_shortcuts import KeyboardShortcutsManager
        
        shortcuts = KeyboardShortcutsManager.DEFAULT_SHORTCUTS
        
        # 有效的修饰键
        valid_modifiers = {'control', 'shift', 'alt', 'meta', 'super'}
        
        for shortcut in shortcuts.keys():
            # 检查格式
            assert shortcut.startswith('<') and shortcut.endswith('>'), \
                f"快捷键格式无效: {shortcut}"
            
            # 解析快捷键
            inner = shortcut[1:-1]
            parts = inner.split('-')
            
            # 至少有一个按键
            assert len(parts) >= 1, f"快捷键缺少按键: {shortcut}"
            
            # 检查修饰键有效性
            for part in parts[:-1]:
                assert part.lower() in valid_modifiers or part.startswith('F'), \
                    f"无效的修饰键 '{part}' 在快捷键 {shortcut} 中"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
