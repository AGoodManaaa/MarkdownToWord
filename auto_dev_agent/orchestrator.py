"""
Orchestrator - 协调整个开发流程的核心组件
"""

from typing import List, Optional, Callable

from .models import (
    Config,
    Task,
    Review,
    ReviewStatus,
    AggregatedFeedback,
    IterationResult,
    DevState,
    TaskResult,
    CodeOutput,
)
from .actor import ActorAgent
from .reviewers import CorrectnessReviewer, StyleReviewer, RobustnessReviewer, BaseReviewer
from .state_manager import StateManager
from .command_executor import CommandExecutor


class Orchestrator:
    """协调整个开发流程的核心组件"""

    def __init__(
        self,
        config: Config,
        output_callback: Optional[Callable[[str], None]] = None,
        confirmation_callback: Optional[Callable[[str], bool]] = None
    ):
        """
        初始化编排器
        
        Args:
            config: 系统配置
            output_callback: 输出回调函数
            confirmation_callback: 危险命令确认回调
        """
        self.config = config
        self.output_callback = output_callback or print
        
        # 初始化组件
        self.actor = ActorAgent(config)
        self.reviewers: List[BaseReviewer] = [
            CorrectnessReviewer(config),
            StyleReviewer(config),
            RobustnessReviewer(config),
        ]
        self.state_manager = StateManager(config.state_dir)
        self.command_executor = CommandExecutor(
            safe_mode=config.safe_mode,
            timeout=config.command_timeout,
            confirmation_callback=confirmation_callback
        )

    def run_task(self, task: Task, resume_state: Optional[DevState] = None) -> TaskResult:
        """
        执行开发任务的主循环
        
        Args:
            task: 要执行的任务
            resume_state: 可选的恢复状态
            
        Returns:
            TaskResult 任务执行结果
        """
        # 初始化或恢复状态
        if resume_state:
            state = resume_state
            self._log(f"恢复任务: {task.id}, 从迭代 {state.current_iteration} 继续")
        else:
            state = DevState(
                task=task,
                current_iteration=0,
                current_code="",
                iteration_history=[],
                status="in_progress"
            )
            self._log(f"开始新任务: {task.id}")
            
            # 生成初始代码
            self._log("Actor 正在生成初始代码...")
            code_output = self.actor.generate_code(task)
            state.current_code = code_output.code
            self._log(f"代码说明: {code_output.explanation}")
            
            # 执行初始命令
            if code_output.commands_to_run:
                self._execute_commands(code_output.commands_to_run)

        # 主循环
        while state.current_iteration < self.config.max_iterations:
            state.current_iteration += 1
            self._log(f"\n=== 迭代 {state.current_iteration}/{self.config.max_iterations} ===")
            
            # 执行单次迭代
            iteration_result = self._run_iteration(state)
            state.iteration_history.append(iteration_result)
            
            # 保存状态
            self.state_manager.save_state(state)
            
            # 检查是否达成共识
            if iteration_result.consensus_reached:
                state.status = "completed"
                self._log("\n✓ 所有 Reviewer 通过！任务完成。")
                break
            
            # 聚合反馈并修改代码
            feedback = self._aggregate_feedback(iteration_result.reviews)
            self._log("\nActor 正在根据反馈修改代码...")
            code_output = self.actor.revise_code(state.current_code, feedback, task)
            state.current_code = code_output.code
            
            # 执行命令
            if code_output.commands_to_run:
                self._execute_commands(code_output.commands_to_run)
        
        # 检查是否达到最大迭代
        if state.status != "completed":
            state.status = "max_iterations_reached"
            self._log(f"\n⚠ 达到最大迭代次数 ({self.config.max_iterations})，输出当前最佳版本。")

        # 保存最终状态
        self.state_manager.save_state(state)
        
        # 构建结果
        return self._build_task_result(state)

    def _run_iteration(self, state: DevState) -> IterationResult:
        """
        执行单次迭代
        
        Args:
            state: 当前开发状态
            
        Returns:
            IterationResult 迭代结果
        """
        self._log(f"\n当前代码:\n```\n{state.current_code[:500]}{'...' if len(state.current_code) > 500 else ''}\n```")
        
        # 收集所有 Reviewer 的审查
        reviews: List[Review] = []
        for reviewer in self.reviewers:
            self._log(f"\n{reviewer.reviewer_type.capitalize()} Reviewer 正在审查...")
            review = reviewer.review(state.current_code, state.task)
            reviews.append(review)
            self._log_review(review)
        
        # 检查共识
        consensus = self._check_consensus(reviews)
        
        return IterationResult(
            iteration_number=state.current_iteration,
            code=state.current_code,
            reviews=reviews,
            consensus_reached=consensus
        )

    def _aggregate_feedback(self, reviews: List[Review]) -> AggregatedFeedback:
        """
        聚合所有Reviewer的反馈
        
        Args:
            reviews: 审查结果列表
            
        Returns:
            AggregatedFeedback 聚合后的反馈
        """
        all_approved = all(r.is_approved for r in reviews)
        combined_issues = []
        combined_suggestions = []
        
        for review in reviews:
            combined_issues.extend(review.issues)
            combined_suggestions.extend(review.suggestions)
        
        # 优先项：分数最低的 Reviewer 的问题
        priority_items = []
        if reviews:
            lowest_score_review = min(reviews, key=lambda r: r.score)
            if lowest_score_review.issues:
                priority_items = lowest_score_review.issues[:3]
        
        return AggregatedFeedback(
            all_approved=all_approved,
            combined_issues=combined_issues,
            combined_suggestions=combined_suggestions,
            priority_items=priority_items
        )

    def _check_consensus(self, reviews: List[Review]) -> bool:
        """
        检查是否达成共识
        
        Args:
            reviews: 审查结果列表
            
        Returns:
            是否所有 Reviewer 都通过
        """
        return all(r.status == ReviewStatus.APPROVED for r in reviews)

    def _execute_commands(self, commands: List[str]) -> None:
        """执行命令列表"""
        for cmd in commands:
            self._log(f"\n执行命令: {cmd}")
            result = self.command_executor.execute(cmd)
            if result.success:
                self._log(f"✓ 成功\n{result.stdout}")
            else:
                self._log(f"✗ 失败\n{result.stderr}")

    def _build_task_result(self, state: DevState) -> TaskResult:
        """构建任务结果"""
        final_reviews = []
        if state.iteration_history:
            final_reviews = state.iteration_history[-1].reviews
        
        summary_parts = [
            f"任务ID: {state.task.id}",
            f"总迭代次数: {state.current_iteration}",
            f"最终状态: {state.status}",
        ]
        
        if final_reviews:
            avg_score = sum(r.score for r in final_reviews) / len(final_reviews)
            summary_parts.append(f"平均评分: {avg_score:.1f}/10")
            
            remaining_issues = []
            for r in final_reviews:
                remaining_issues.extend(r.issues)
            if remaining_issues:
                summary_parts.append(f"剩余问题: {len(remaining_issues)}")
        
        return TaskResult(
            success=state.status == "completed",
            final_code=state.current_code,
            total_iterations=state.current_iteration,
            final_reviews=final_reviews,
            summary="\n".join(summary_parts)
        )

    def _log(self, message: str) -> None:
        """输出日志"""
        if self.output_callback:
            self.output_callback(message)

    def _log_review(self, review: Review) -> None:
        """输出审查结果"""
        status_icon = "✓" if review.is_approved else "✗"
        self._log(f"  {status_icon} 状态: {review.status.value}, 评分: {review.score}/10")
        if review.issues:
            self._log(f"  问题: {', '.join(review.issues[:3])}")
        if review.suggestions:
            self._log(f"  建议: {', '.join(review.suggestions[:3])}")
