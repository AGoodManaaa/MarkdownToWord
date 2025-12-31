# -*- coding: utf-8 -*-
"""
Tests for KeyboardShortcutsManager module.

Includes both unit tests and property-based tests.
"""

import unittest
from unittest.mock import MagicMock, patch
from hypothesis import given, strategies as st, settings

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ui.keyboard_shortcuts import KeyboardShortcutsManager


class MockApp:
    """Mock application for testing."""
    
    def __init__(self):
        self._bindings = {}
        self.open_file = MagicMock()
        self.save_file = MagicMock()
        self.toggle_preview = MagicMock()
        self.focus_mode = MagicMock()
        self.focus_mode.toggle = MagicMock()
        self.command_palette = MagicMock()
        self.command_palette.show = MagicMock()
    
    def bind(self, key, handler):
        self._bindings[key] = handler
    
    def unbind(self, key):
        self._bindings.pop(key, None)
    
    def change_font_size(self, delta):
        pass


class TestKeyboardShortcutsManagerUnit(unittest.TestCase):
    """Unit tests for KeyboardShortcutsManager."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.app = MockApp()
        self.manager = KeyboardShortcutsManager(self.app)
    
    def test_load_default_shortcuts(self):
        """Test that default shortcuts are loaded."""
        self.manager.load_from_config()
        
        self.assertIn('<Control-o>', self.manager.shortcuts)
        self.assertIn('<Control-s>', self.manager.shortcuts)
        self.assertEqual(self.manager.shortcuts['<Control-o>'], 'open_file')
    
    def test_load_custom_shortcuts_override(self):
        """Test that custom shortcuts override defaults."""
        config = {
            'shortcuts': {
                '<Control-o>': 'custom_action',
                '<Control-n>': 'new_action'
            }
        }
        
        self.manager.load_from_config(config)
        
        self.assertEqual(self.manager.shortcuts['<Control-o>'], 'custom_action')
        self.assertEqual(self.manager.shortcuts['<Control-n>'], 'new_action')
        # Default shortcuts should still exist
        self.assertIn('<Control-s>', self.manager.shortcuts)
    
    def test_bind_all_creates_bindings(self):
        """Test that bind_all creates bindings on the app."""
        self.manager.load_from_config()
        self.manager.bind_all()
        
        # Check that bindings were created
        self.assertIn('<Control-o>', self.app._bindings)
        self.assertIn('<Control-s>', self.app._bindings)
    
    def test_resolve_simple_action(self):
        """Test resolving a simple action string."""
        handler = self.manager._resolve_action('open_file')
        
        self.assertIsNotNone(handler)
        self.assertEqual(handler, self.app.open_file)
    
    def test_resolve_nested_action(self):
        """Test resolving a nested action string."""
        handler = self.manager._resolve_action('focus_mode.toggle')
        
        self.assertIsNotNone(handler)
        self.assertEqual(handler, self.app.focus_mode.toggle)
    
    def test_resolve_nonexistent_action(self):
        """Test resolving a nonexistent action returns None."""
        handler = self.manager._resolve_action('nonexistent_method')
        
        self.assertIsNone(handler)
    
    def test_get_shortcut_for_action(self):
        """Test getting shortcut for a specific action."""
        self.manager.load_from_config()
        
        shortcut = self.manager.get_shortcut_for_action('open_file')
        
        self.assertEqual(shortcut, '<Control-o>')
    
    def test_get_all_shortcuts(self):
        """Test getting all shortcuts returns a copy."""
        self.manager.load_from_config()
        
        shortcuts = self.manager.get_all_shortcuts()
        shortcuts['<Control-x>'] = 'test'
        
        self.assertNotIn('<Control-x>', self.manager.shortcuts)
    
    def test_unbind_all(self):
        """Test unbinding all shortcuts."""
        self.manager.load_from_config()
        self.manager.bind_all()
        
        self.manager.unbind_all()
        
        self.assertEqual(len(self.manager._bound_handlers), 0)


class TestKeyboardShortcutsManagerProperty(unittest.TestCase):
    """
    Property-based tests for KeyboardShortcutsManager.
    
    **Feature: code-optimization, Property 2: 快捷键配置一致性**
    **Validates: Requirements 3.3**
    """
    
    def setUp(self):
        """Set up test fixtures."""
        self.app = MockApp()
    
    @given(st.dictionaries(
        keys=st.text(min_size=1, max_size=20).map(lambda x: f'<Control-{x}>'),
        values=st.sampled_from(['open_file', 'save_file', 'toggle_preview']),
        min_size=0,
        max_size=10
    ))
    @settings(max_examples=100)
    def test_shortcut_config_consistency(self, custom_shortcuts):
        """
        Property 2: 快捷键配置一致性
        
        *For any* 快捷键配置，加载后绑定的快捷键应该与配置完全一致。
        
        **Feature: code-optimization, Property 2: 快捷键配置一致性**
        **Validates: Requirements 3.3**
        """
        manager = KeyboardShortcutsManager(self.app)
        config = {'shortcuts': custom_shortcuts}
        
        manager.load_from_config(config)
        
        # Property: all custom shortcuts should be in the loaded shortcuts
        for key, action in custom_shortcuts.items():
            self.assertIn(key, manager.shortcuts,
                f"Custom shortcut '{key}' should be loaded")
            self.assertEqual(manager.shortcuts[key], action,
                f"Shortcut '{key}' should map to '{action}'")
    
    @given(st.lists(
        st.tuples(
            st.text(min_size=1, max_size=10).map(lambda x: f'<Control-{x}>'),
            st.sampled_from(['open_file', 'save_file', 'toggle_preview'])
        ),
        min_size=0,
        max_size=10,
        unique_by=lambda x: x[0]
    ))
    @settings(max_examples=100)
    def test_bound_shortcuts_match_config(self, shortcut_pairs):
        """
        Property: After bind_all(), bound handlers match the configuration.
        
        **Feature: code-optimization, Property 2: 快捷键配置一致性**
        **Validates: Requirements 3.3**
        """
        manager = KeyboardShortcutsManager(self.app)
        custom_shortcuts = dict(shortcut_pairs)
        config = {'shortcuts': custom_shortcuts}
        
        manager.load_from_config(config)
        manager.bind_all()
        
        # Property: number of bound handlers should match resolvable shortcuts
        resolvable_count = sum(
            1 for action in manager.shortcuts.values()
            if manager._resolve_action(action) is not None
        )
        
        self.assertEqual(len(manager._bound_handlers), resolvable_count,
            "Number of bound handlers should match resolvable shortcuts")


if __name__ == '__main__':
    unittest.main()
