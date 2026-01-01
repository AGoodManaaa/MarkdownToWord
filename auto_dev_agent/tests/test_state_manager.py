"""
Property-based tests for StateManager
"""

import pytest
import tempfile
import shutil
from pathlib import Path
from hypothesis import given, strategies as st, settings

from auto_dev_agent.models import DevState, Task, IterationResult, Review, ReviewStatus
from auto_dev_agent.state_manager import StateManager


# Reuse strategies from test_models
task_strategy = st.builds(
    Task,
    id=st.text(min_size=1, max_size=10, alphabet=st.characters(whitelist_categories=("L", "N"))),
    description=st.text(min_size=1, max_size=200),
    context=st.one_of(st.none(), st.text(max_size=100)),
    created_at=st.text(min_size=1, max_size=30),
)

dev_state_strategy = st.builds(
    DevState,
    task=task_strategy,
    current_iteration=st.integers(min_value=0, max_value=100),
    current_code=st.text(max_size=500),
    iteration_history=st.just([]),
    status=st.sampled_from(["in_progress", "completed", "max_iterations_reached"]),
)


class TestStatePersistence:
    """
    **Feature: auto-dev-agent, Property 15: State Persistence After Iteration**
    **Validates: Requirements 7.1**
    
    For any completed iteration, the state file SHALL exist and contain the current state.
    """

    @pytest.fixture
    def temp_state_dir(self):
        """Create a temporary directory for state files"""
        temp_dir = tempfile.mkdtemp()
        yield temp_dir
        shutil.rmtree(temp_dir)

    @given(dev_state_strategy)
    @settings(max_examples=100)
    def test_state_persistence_after_save(self, state: DevState):
        """Test that state file exists after save"""
        with tempfile.TemporaryDirectory() as temp_dir:
            manager = StateManager(state_dir=temp_dir)
            manager.save_state(state)
            
            # Verify file exists
            assert manager.state_exists(state.task.id)
            
            # Verify content can be loaded
            loaded = manager.load_state(state.task.id)
            assert loaded is not None
            assert loaded.task.id == state.task.id
            assert loaded.current_code == state.current_code

    @given(dev_state_strategy)
    @settings(max_examples=50)
    def test_save_load_round_trip(self, state: DevState):
        """Test save then load produces equivalent state"""
        with tempfile.TemporaryDirectory() as temp_dir:
            manager = StateManager(state_dir=temp_dir)
            manager.save_state(state)
            loaded = manager.load_state(state.task.id)
            
            assert loaded.task.id == state.task.id
            assert loaded.task.description == state.task.description
            assert loaded.current_iteration == state.current_iteration
            assert loaded.current_code == state.current_code
            assert loaded.status == state.status


class TestStateRecoveryDetection:
    """
    **Feature: auto-dev-agent, Property 16: State Recovery Detection**
    **Validates: Requirements 7.2**
    
    For any existing state file, the system SHALL detect it on startup.
    """

    @given(st.lists(dev_state_strategy, min_size=1, max_size=5, unique_by=lambda s: s.task.id))
    @settings(max_examples=50)
    def test_detect_existing_states(self, states: list):
        """Test that existing state files are detected"""
        with tempfile.TemporaryDirectory() as temp_dir:
            manager = StateManager(state_dir=temp_dir)
            
            # Save all states
            for state in states:
                manager.save_state(state)
            
            # Verify all can be detected
            all_ids = manager.get_all_task_ids()
            for state in states:
                assert state.task.id in all_ids

    @given(st.lists(dev_state_strategy.filter(lambda s: s.status == "in_progress"), 
                    min_size=1, max_size=5, unique_by=lambda s: s.task.id))
    @settings(max_examples=50)
    def test_list_pending_tasks(self, states: list):
        """Test that pending tasks are correctly listed"""
        with tempfile.TemporaryDirectory() as temp_dir:
            manager = StateManager(state_dir=temp_dir)
            
            for state in states:
                manager.save_state(state)
            
            pending = manager.list_pending_tasks()
            for state in states:
                assert state.task.id in pending

    def test_empty_dir_returns_empty_list(self):
        """Test that empty directory returns empty list"""
        with tempfile.TemporaryDirectory() as temp_dir:
            manager = StateManager(state_dir=temp_dir)
            assert manager.list_pending_tasks() == []
            assert manager.get_all_task_ids() == []
