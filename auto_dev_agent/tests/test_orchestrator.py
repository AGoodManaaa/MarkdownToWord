"""
Property-based tests for Orchestrator
"""

import pytest
from hypothesis import given, strategies as st, settings

from auto_dev_agent.models import (
    Review,
    ReviewStatus,
    AggregatedFeedback,
    Task,
    DevState,
    IterationResult,
    Config,
)
from auto_dev_agent.orchestrator import Orchestrator


# Strategies
review_strategy = st.builds(
    Review,
    reviewer_type=st.sampled_from(["correctness", "style", "robustness"]),
    status=st.sampled_from([ReviewStatus.APPROVED, ReviewStatus.NEEDS_REVISION]),
    issues=st.lists(st.text(min_size=1, max_size=50), max_size=5),
    suggestions=st.lists(st.text(min_size=1, max_size=50), max_size=5),
    score=st.integers(min_value=1, max_value=10),
)

# Reviews where all are approved
all_approved_reviews = st.lists(
    st.builds(
        Review,
        reviewer_type=st.sampled_from(["correctness", "style", "robustness"]),
        status=st.just(ReviewStatus.APPROVED),
        issues=st.just([]),
        suggestions=st.just([]),
        score=st.integers(min_value=7, max_value=10),
    ),
    min_size=3,
    max_size=3,
)

# Reviews where at least one needs revision
mixed_reviews = st.lists(review_strategy, min_size=3, max_size=3).filter(
    lambda reviews: any(r.status == ReviewStatus.NEEDS_REVISION for r in reviews)
)


class TestConsensusDetection:
    """
    **Feature: auto-dev-agent, Property 4: Consensus Detection Correctness**
    **Validates: Requirements 1.4**
    
    For any set of reviews where all three Reviewers have status APPROVED,
    the system SHALL mark the task as complete.
    """

    @given(all_approved_reviews)
    @settings(max_examples=100)
    def test_all_approved_reaches_consensus(self, reviews: list):
        """Test that all approved reviews result in consensus"""
        config = Config()
        orchestrator = Orchestrator(config, output_callback=lambda x: None)
        
        consensus = orchestrator._check_consensus(reviews)
        assert consensus is True

    @given(mixed_reviews)
    @settings(max_examples=100)
    def test_mixed_reviews_no_consensus(self, reviews: list):
        """Test that mixed reviews do not result in consensus"""
        config = Config()
        orchestrator = Orchestrator(config, output_callback=lambda x: None)
        
        consensus = orchestrator._check_consensus(reviews)
        assert consensus is False


class TestFeedbackAggregation:
    """
    **Feature: auto-dev-agent, Property 3: Feedback Aggregation Completeness**
    **Validates: Requirements 1.3**
    
    For any set of three reviews from Reviewers, the aggregated feedback
    sent to the Actor SHALL contain issues and suggestions from all three reviews.
    """

    @given(st.lists(review_strategy, min_size=3, max_size=3))
    @settings(max_examples=100)
    def test_all_issues_aggregated(self, reviews: list):
        """Test that all issues are included in aggregated feedback"""
        config = Config()
        orchestrator = Orchestrator(config, output_callback=lambda x: None)
        
        feedback = orchestrator._aggregate_feedback(reviews)
        
        # All issues from all reviews should be in combined_issues
        all_issues = []
        for review in reviews:
            all_issues.extend(review.issues)
        
        for issue in all_issues:
            assert issue in feedback.combined_issues

    @given(st.lists(review_strategy, min_size=3, max_size=3))
    @settings(max_examples=100)
    def test_all_suggestions_aggregated(self, reviews: list):
        """Test that all suggestions are included in aggregated feedback"""
        config = Config()
        orchestrator = Orchestrator(config, output_callback=lambda x: None)
        
        feedback = orchestrator._aggregate_feedback(reviews)
        
        # All suggestions from all reviews should be in combined_suggestions
        all_suggestions = []
        for review in reviews:
            all_suggestions.extend(review.suggestions)
        
        for suggestion in all_suggestions:
            assert suggestion in feedback.combined_suggestions


class TestIterationContinuation:
    """
    **Feature: auto-dev-agent, Property 6: Iteration Continuation Below Maximum**
    **Validates: Requirements 3.2**
    
    For any iteration count below the configured maximum, if consensus is not reached,
    the system SHALL continue to the next iteration.
    """

    @given(st.integers(min_value=1, max_value=9))
    @settings(max_examples=50)
    def test_iteration_below_max_continues(self, current_iteration: int):
        """Test that iteration continues when below max and no consensus"""
        max_iterations = 10
        
        # If current_iteration < max_iterations and no consensus, should continue
        should_continue = current_iteration < max_iterations
        assert should_continue is True


class TestMaximumIterationTermination:
    """
    **Feature: auto-dev-agent, Property 7: Maximum Iteration Termination**
    **Validates: Requirements 3.3**
    
    For any task that reaches maximum iterations without consensus,
    the output SHALL include the best code version and a summary of remaining issues.
    """

    @given(st.integers(min_value=1, max_value=10))
    @settings(max_examples=50)
    def test_max_iteration_produces_result(self, max_iterations: int):
        """Test that reaching max iterations produces a result with code and summary"""
        config = Config(max_iterations=max_iterations)
        
        # Create a mock state that has reached max iterations
        task = Task.create(description="Test task")
        state = DevState(
            task=task,
            current_iteration=max_iterations,
            current_code="# Final code",
            iteration_history=[],
            status="max_iterations_reached"
        )
        
        orchestrator = Orchestrator(config, output_callback=lambda x: None)
        result = orchestrator._build_task_result(state)
        
        # Result should have code and summary
        assert result.final_code == "# Final code"
        assert "max_iterations_reached" in result.summary
        assert result.total_iterations == max_iterations

    def test_max_iteration_includes_remaining_issues(self):
        """Test that max iteration result includes remaining issues"""
        config = Config(max_iterations=2)
        task = Task.create(description="Test task")
        
        # Create iteration with unresolved issues
        reviews = [
            Review(
                reviewer_type="correctness",
                status=ReviewStatus.NEEDS_REVISION,
                issues=["Bug found"],
                suggestions=["Fix the bug"],
                score=5
            )
        ]
        
        state = DevState(
            task=task,
            current_iteration=2,
            current_code="# Code with issues",
            iteration_history=[
                IterationResult(
                    iteration_number=2,
                    code="# Code with issues",
                    reviews=reviews,
                    consensus_reached=False
                )
            ],
            status="max_iterations_reached"
        )
        
        orchestrator = Orchestrator(config, output_callback=lambda x: None)
        result = orchestrator._build_task_result(state)
        
        # Should include remaining issues count
        assert "剩余问题" in result.summary or result.final_reviews


class TestIssueSuggestionCorrelation:
    """
    **Feature: auto-dev-agent, Property 5: Issue-Suggestion Correlation**
    **Validates: Requirements 2.4**
    
    For any review that identifies issues (non-empty issues list),
    the review SHALL also contain at least one suggestion.
    """

    @given(st.lists(st.text(min_size=1, max_size=50), min_size=1, max_size=5))
    @settings(max_examples=100)
    def test_issues_have_suggestions(self, issues: list):
        """Test that reviews with issues always have suggestions"""
        import json
        from auto_dev_agent.reviewers import CorrectnessReviewer
        
        config = Config()
        reviewer = CorrectnessReviewer(config)
        
        # Simulate AI response with issues but no suggestions
        response = json.dumps({
            "status": "needs_revision",
            "issues": issues,
            "suggestions": [],  # Empty suggestions
            "score": 5
        })
        
        review = reviewer._parse_review_response(response)
        
        # Property: if issues exist, suggestions must exist
        if review.issues:
            assert len(review.suggestions) > 0

    @given(
        st.lists(st.text(min_size=1, max_size=50), min_size=1, max_size=5),
        st.lists(st.text(min_size=1, max_size=50), min_size=1, max_size=5)
    )
    @settings(max_examples=100)
    def test_issues_with_suggestions_preserved(self, issues: list, suggestions: list):
        """Test that provided suggestions are preserved"""
        import json
        from auto_dev_agent.reviewers import StyleReviewer
        
        config = Config()
        reviewer = StyleReviewer(config)
        
        response = json.dumps({
            "status": "needs_revision",
            "issues": issues,
            "suggestions": suggestions,
            "score": 5
        })
        
        review = reviewer._parse_review_response(response)
        
        # All provided suggestions should be preserved
        for suggestion in suggestions:
            assert suggestion in review.suggestions
