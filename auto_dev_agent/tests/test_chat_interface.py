"""
Property-based tests for Chat Interface
"""

import pytest
from hypothesis import given, strategies as st, settings

from auto_dev_agent.models import (
    Review,
    ReviewStatus,
    IterationResult,
    CodeOutput,
)
from auto_dev_agent.chat_interface import MessageFormatter, AiideChatInterface


# Strategies
review_strategy = st.builds(
    Review,
    reviewer_type=st.sampled_from(["correctness", "style", "robustness"]),
    status=st.sampled_from([ReviewStatus.APPROVED, ReviewStatus.NEEDS_REVISION]),
    issues=st.lists(st.text(min_size=1, max_size=50), max_size=5),
    suggestions=st.lists(st.text(min_size=1, max_size=50), max_size=5),
    score=st.integers(min_value=1, max_value=10),
)

iteration_result_strategy = st.builds(
    IterationResult,
    iteration_number=st.integers(min_value=1, max_value=100),
    code=st.text(max_size=500),
    reviews=st.lists(review_strategy, min_size=1, max_size=3),
    consensus_reached=st.booleans(),
)

code_output_strategy = st.builds(
    CodeOutput,
    code=st.text(min_size=1, max_size=500),
    explanation=st.text(max_size=200),
    commands_to_run=st.lists(st.text(min_size=1, max_size=50), max_size=3),
)


class TestOutputCompleteness:
    """
    **Feature: auto-dev-agent, Property 8: Output Completeness**
    **Validates: Requirements 4.1, 4.2, 4.3**
    
    For any iteration, the system output SHALL include the iteration number,
    code changes, and all Reviewer comments.
    """

    @given(iteration_result_strategy)
    @settings(max_examples=100)
    def test_iteration_summary_contains_iteration_number(self, result: IterationResult):
        """Test that iteration summary contains iteration number"""
        formatter = MessageFormatter()
        message = formatter.format_iteration_summary(result)
        
        assert str(result.iteration_number) in message.content
        assert message.metadata["iteration"] == result.iteration_number

    @given(review_strategy)
    @settings(max_examples=100)
    def test_review_format_contains_all_info(self, review: Review):
        """Test that formatted review contains all information"""
        formatter = MessageFormatter()
        message = formatter.format_review(review)
        
        # Should contain reviewer type
        assert review.reviewer_type.capitalize() in message.content
        # Should contain status
        assert review.status.value in message.content
        # Should contain score
        assert str(review.score) in message.content

    @given(code_output_strategy)
    @settings(max_examples=100)
    def test_code_output_format_contains_code(self, output: CodeOutput):
        """Test that formatted code output contains the code"""
        formatter = MessageFormatter()
        message = formatter.format_code_output(output)
        
        # Should contain the code
        assert output.code in message.content
        # Should contain explanation
        assert output.explanation in message.content

    @given(iteration_result_strategy)
    @settings(max_examples=50)
    def test_all_reviews_in_summary(self, result: IterationResult):
        """Test that all reviewer statuses are shown in summary"""
        formatter = MessageFormatter()
        message = formatter.format_iteration_summary(result)
        
        # Each reviewer type should be mentioned
        for review in result.reviews:
            assert review.reviewer_type.capitalize() in message.content
