"""
Property-based tests for data models
**Feature: auto-dev-agent, Property 14: State Serialization Round-Trip**
"""

import pytest
from hypothesis import given, strategies as st, settings

from auto_dev_agent.models import (
    Config,
    Task,
    CodeOutput,
    Review,
    ReviewStatus,
    AggregatedFeedback,
    IterationResult,
    DevState,
    CommandResult,
)


# Strategies for generating test data
review_status_strategy = st.sampled_from([ReviewStatus.APPROVED, ReviewStatus.NEEDS_REVISION])

task_strategy = st.builds(
    Task,
    id=st.text(min_size=1, max_size=10, alphabet=st.characters(whitelist_categories=("L", "N"))),
    description=st.text(min_size=1, max_size=200),
    context=st.one_of(st.none(), st.text(max_size=100)),
    created_at=st.text(min_size=1, max_size=30),
)

command_result_strategy = st.builds(
    CommandResult,
    command=st.text(min_size=1, max_size=50),
    success=st.booleans(),
    stdout=st.text(max_size=100),
    stderr=st.text(max_size=100),
    return_code=st.integers(min_value=-128, max_value=127),
)

review_strategy = st.builds(
    Review,
    reviewer_type=st.sampled_from(["correctness", "style", "robustness"]),
    status=review_status_strategy,
    issues=st.lists(st.text(min_size=1, max_size=50), max_size=5),
    suggestions=st.lists(st.text(min_size=1, max_size=50), max_size=5),
    score=st.integers(min_value=1, max_value=10),
)

iteration_result_strategy = st.builds(
    IterationResult,
    iteration_number=st.integers(min_value=1, max_value=100),
    code=st.text(max_size=500),
    reviews=st.lists(review_strategy, max_size=3),
    consensus_reached=st.booleans(),
    command_results=st.lists(command_result_strategy, max_size=3),
)

dev_state_strategy = st.builds(
    DevState,
    task=task_strategy,
    current_iteration=st.integers(min_value=0, max_value=100),
    current_code=st.text(max_size=500),
    iteration_history=st.lists(iteration_result_strategy, max_size=5),
    status=st.sampled_from(["in_progress", "completed", "max_iterations_reached"]),
)


class TestStateSerializationRoundTrip:
    """
    **Feature: auto-dev-agent, Property 14: State Serialization Round-Trip**
    **Validates: Requirements 7.3, 7.4**
    
    For any DevState, serializing to JSON and then deserializing 
    SHALL produce an equivalent DevState object.
    """

    @given(dev_state_strategy)
    @settings(max_examples=100)
    def test_dev_state_round_trip(self, state: DevState):
        """Test DevState serialization round-trip"""
        json_str = state.to_json()
        restored = DevState.from_json(json_str)
        
        assert restored.task.id == state.task.id
        assert restored.task.description == state.task.description
        assert restored.current_iteration == state.current_iteration
        assert restored.current_code == state.current_code
        assert restored.status == state.status
        assert len(restored.iteration_history) == len(state.iteration_history)

    @given(task_strategy)
    @settings(max_examples=100)
    def test_task_round_trip(self, task: Task):
        """Test Task serialization round-trip"""
        d = task.to_dict()
        restored = Task.from_dict(d)
        
        assert restored.id == task.id
        assert restored.description == task.description
        assert restored.context == task.context

    @given(review_strategy)
    @settings(max_examples=100)
    def test_review_round_trip(self, review: Review):
        """Test Review serialization round-trip"""
        d = review.to_dict()
        restored = Review.from_dict(d)
        
        assert restored.reviewer_type == review.reviewer_type
        assert restored.status == review.status
        assert restored.issues == review.issues
        assert restored.suggestions == review.suggestions
        assert restored.score == review.score

    @given(iteration_result_strategy)
    @settings(max_examples=100)
    def test_iteration_result_round_trip(self, result: IterationResult):
        """Test IterationResult serialization round-trip"""
        d = result.to_dict()
        restored = IterationResult.from_dict(d)
        
        assert restored.iteration_number == result.iteration_number
        assert restored.code == result.code
        assert restored.consensus_reached == result.consensus_reached
        assert len(restored.reviews) == len(result.reviews)
