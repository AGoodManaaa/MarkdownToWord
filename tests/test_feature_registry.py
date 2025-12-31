# -*- coding: utf-8 -*-
"""
Tests for FeatureRegistry module.

Includes both unit tests and property-based tests.
"""

import unittest
import logging
from unittest.mock import MagicMock, patch
from hypothesis import given, strategies as st, settings

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ui.feature_registry import FeatureRegistry


class MockApp:
    """Mock application for testing."""
    pass


class MockFeature:
    """Mock feature class for testing."""
    instance_count = 0
    
    def __init__(self, app, *args, **kwargs):
        self.app = app
        self.args = args
        self.kwargs = kwargs
        MockFeature.instance_count += 1
        self.instance_id = MockFeature.instance_count


class MockFeatureA(MockFeature):
    """Another mock feature class."""
    pass


class MockFeatureB(MockFeature):
    """Yet another mock feature class."""
    pass


class TestFeatureRegistryUnit(unittest.TestCase):
    """Unit tests for FeatureRegistry."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.app = MockApp()
        self.registry = FeatureRegistry(self.app)
        MockFeature.instance_count = 0
    
    def test_register_creates_instance(self):
        """Test that register() creates a new feature instance."""
        feature = self.registry.register('test_feature', MockFeature)
        
        self.assertIsInstance(feature, MockFeature)
        self.assertEqual(feature.app, self.app)
    
    def test_register_returns_same_instance_on_duplicate(self):
        """Test that registering the same name twice returns the existing instance.
        
        _Requirements: 1.1, 1.3_
        """
        feature1 = self.registry.register('test_feature', MockFeature)
        feature2 = self.registry.register('test_feature', MockFeature)
        
        self.assertIs(feature1, feature2)
        self.assertEqual(feature1.instance_id, feature2.instance_id)
    
    def test_register_logs_warning_on_duplicate(self):
        """Test that duplicate registration logs a warning.
        
        _Requirements: 1.1, 1.3_
        """
        with patch.object(logging.getLogger('ui.feature_registry'), 'warning') as mock_warning:
            self.registry.register('test_feature', MockFeature)
            self.registry.register('test_feature', MockFeature)
            
            mock_warning.assert_called_once()
            self.assertIn('already registered', mock_warning.call_args[0][0])
    
    def test_get_returns_registered_feature(self):
        """Test that get() returns the correct registered feature.
        
        _Requirements: 1.1, 1.3_
        """
        registered = self.registry.register('my_feature', MockFeature)
        retrieved = self.registry.get('my_feature')
        
        self.assertIs(registered, retrieved)
    
    def test_get_returns_none_for_unregistered(self):
        """Test that get() returns None for unregistered features."""
        result = self.registry.get('nonexistent')
        
        self.assertIsNone(result)
    
    def test_has_returns_true_for_registered(self):
        """Test that has() returns True for registered features."""
        self.registry.register('test_feature', MockFeature)
        
        self.assertTrue(self.registry.has('test_feature'))
    
    def test_has_returns_false_for_unregistered(self):
        """Test that has() returns False for unregistered features."""
        self.assertFalse(self.registry.has('nonexistent'))
    
    def test_count_returns_correct_number(self):
        """Test that count() returns the correct number of features."""
        self.assertEqual(self.registry.count(), 0)
        
        self.registry.register('feature1', MockFeature)
        self.assertEqual(self.registry.count(), 1)
        
        self.registry.register('feature2', MockFeature)
        self.assertEqual(self.registry.count(), 2)
    
    def test_get_all_returns_copy(self):
        """Test that get_all() returns a copy of the features dict."""
        self.registry.register('feature1', MockFeature)
        
        all_features = self.registry.get_all()
        all_features['feature2'] = 'fake'
        
        self.assertFalse(self.registry.has('feature2'))
    
    def test_initialize_all_registers_multiple_features(self):
        """Test that initialize_all() registers multiple features."""
        definitions = [
            ('feature_a', MockFeatureA),
            ('feature_b', MockFeatureB),
        ]
        
        self.registry.initialize_all(definitions)
        
        self.assertTrue(self.registry.has('feature_a'))
        self.assertTrue(self.registry.has('feature_b'))
        self.assertIsInstance(self.registry.get('feature_a'), MockFeatureA)
        self.assertIsInstance(self.registry.get('feature_b'), MockFeatureB)
    
    def test_initialize_all_only_runs_once(self):
        """Test that initialize_all() only runs once."""
        definitions = [('feature_a', MockFeatureA)]
        
        self.registry.initialize_all(definitions)
        initial_count = MockFeature.instance_count
        
        # Try to initialize again
        self.registry.initialize_all([('feature_b', MockFeatureB)])
        
        # Should not have created new instances
        self.assertEqual(MockFeature.instance_count, initial_count)
        self.assertFalse(self.registry.has('feature_b'))
    
    def test_register_with_args_and_kwargs(self):
        """Test that register() passes args and kwargs to feature constructor."""
        feature = self.registry.register('test', MockFeature, 'arg1', 'arg2', key='value')
        
        self.assertEqual(feature.args, ('arg1', 'arg2'))
        self.assertEqual(feature.kwargs, {'key': 'value'})


class TestFeatureRegistryProperty(unittest.TestCase):
    """Property-based tests for FeatureRegistry.
    
    **Feature: code-optimization, Property 1: Feature 单例性**
    **Validates: Requirements 1.1**
    """
    
    def setUp(self):
        """Set up test fixtures."""
        self.app = MockApp()
        MockFeature.instance_count = 0
    
    @given(st.lists(st.text(min_size=1, max_size=20, alphabet=st.characters(whitelist_categories=('L', 'N'))), min_size=1, max_size=10))
    @settings(max_examples=100)
    def test_feature_singleton_property(self, feature_names):
        """
        Property 1: Feature 单例性
        
        *For any* Feature 类型，在 App 生命周期内，FeatureRegistry 应该只创建一个实例。
        
        **Feature: code-optimization, Property 1: Feature 单例性**
        **Validates: Requirements 1.1**
        """
        registry = FeatureRegistry(self.app)
        MockFeature.instance_count = 0
        
        # Register each feature name multiple times
        instances = {}
        for name in feature_names:
            # First registration
            instance1 = registry.register(name, MockFeature)
            instances[name] = instance1
            
            # Second registration (should return same instance)
            instance2 = registry.register(name, MockFeature)
            
            # Property: same name always returns same instance
            self.assertIs(instance1, instance2, 
                f"Feature '{name}' should return the same instance on duplicate registration")
        
        # Property: number of unique instances equals number of unique names
        unique_names = set(feature_names)
        self.assertEqual(
            MockFeature.instance_count, 
            len(unique_names),
            f"Should create exactly {len(unique_names)} instances for {len(unique_names)} unique names"
        )
    
    @given(st.lists(st.text(min_size=1, max_size=10, alphabet=st.characters(whitelist_categories=('L', 'N'))), min_size=0, max_size=20, unique=True))
    @settings(max_examples=100)
    def test_get_returns_registered_instance_property(self, feature_names):
        """
        Property: get() always returns the same instance that was registered.
        
        **Feature: code-optimization, Property 1: Feature 单例性**
        **Validates: Requirements 1.1**
        """
        registry = FeatureRegistry(self.app)
        
        # Register all features
        registered_instances = {}
        for name in feature_names:
            instance = registry.register(name, MockFeature)
            registered_instances[name] = instance
        
        # Property: get() returns the exact same instance
        for name, expected_instance in registered_instances.items():
            retrieved = registry.get(name)
            self.assertIs(retrieved, expected_instance,
                f"get('{name}') should return the same instance that was registered")
    
    @given(st.lists(st.text(min_size=1, max_size=10, alphabet=st.characters(whitelist_categories=('L', 'N'))), min_size=0, max_size=20, unique=True))
    @settings(max_examples=100)
    def test_count_equals_unique_registrations(self, feature_names):
        """
        Property: count() equals the number of unique feature names registered.
        
        **Feature: code-optimization, Property 1: Feature 单例性**
        **Validates: Requirements 1.1**
        """
        registry = FeatureRegistry(self.app)
        
        for name in feature_names:
            registry.register(name, MockFeature)
            # Register again (should not increase count)
            registry.register(name, MockFeature)
        
        self.assertEqual(registry.count(), len(feature_names),
            f"count() should equal {len(feature_names)} unique registrations")


if __name__ == '__main__':
    unittest.main()
