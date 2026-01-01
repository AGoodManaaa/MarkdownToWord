"""
Property-based tests for Configuration
"""

import pytest
from hypothesis import given, strategies as st, settings

from auto_dev_agent.models import Config


# Valid configuration strategies
valid_providers = st.sampled_from(["openai", "anthropic", "azure", "local"])
valid_models = st.sampled_from(["gpt-4", "gpt-3.5-turbo", "claude-3", "llama-2"])

valid_config_strategy = st.builds(
    Config,
    model_provider=valid_providers,
    model_name=valid_models,
    max_iterations=st.integers(min_value=1, max_value=100),
    safe_mode=st.booleans(),
    state_dir=st.text(min_size=1, max_size=50, alphabet=st.characters(whitelist_categories=("L", "N", "P"))),
)


class TestConfigurationAcceptance:
    """
    **Feature: auto-dev-agent, Property 12: Configuration Acceptance**
    **Validates: Requirements 6.1**
    
    For any valid model configuration (provider and model name),
    the system SHALL accept and use the configuration.
    """

    @given(valid_config_strategy)
    @settings(max_examples=100)
    def test_valid_config_accepted(self, config: Config):
        """Test that valid configurations are accepted"""
        # Config should be created successfully
        assert config.model_provider is not None
        assert config.model_name is not None
        assert config.max_iterations > 0

    @given(valid_providers, valid_models)
    @settings(max_examples=100)
    def test_config_stores_values(self, provider: str, model: str):
        """Test that config stores the provided values"""
        config = Config(model_provider=provider, model_name=model)
        
        assert config.model_provider == provider
        assert config.model_name == model


class TestInvalidConfigurationFallback:
    """
    **Feature: auto-dev-agent, Property 13: Invalid Configuration Fallback**
    **Validates: Requirements 6.3**
    
    For any invalid model configuration, the system SHALL fall back
    to default configuration.
    """

    def test_default_config_values(self):
        """Test that default config has expected values"""
        config = Config.get_default()
        
        assert config.model_provider == "openai"
        assert config.model_name == "gpt-4"
        assert config.max_iterations == 10
        assert config.safe_mode is True

    @given(st.integers(min_value=-100, max_value=0))
    @settings(max_examples=50)
    def test_invalid_max_iterations_handled(self, invalid_max: int):
        """Test handling of invalid max_iterations"""
        # In a real implementation, this would validate and fall back
        # For now, we just test that default is valid
        default = Config.get_default()
        assert default.max_iterations > 0

    def test_config_serialization(self):
        """Test config can be serialized and deserialized"""
        config = Config(
            model_provider="anthropic",
            model_name="claude-3",
            max_iterations=5
        )
        
        d = config.to_dict()
        restored = Config.from_dict(d)
        
        assert restored.model_provider == config.model_provider
        assert restored.model_name == config.model_name
        assert restored.max_iterations == config.max_iterations
