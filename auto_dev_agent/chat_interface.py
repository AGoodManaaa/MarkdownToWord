"""
AIIDE Chat Interface - 与AIIDE侧边栏集成的接口
"""

from typing import List, Optional, Dict, Any

from .models import (
    ChatMessage,
    CodeOutput,
    Review,
    IterationResult,
    ReviewStatus,
)


class MessageFormatter:
    """消息格式化器"""

    def format_code_output(self, output: CodeOutput) -> ChatMessage:
        """
        格式化代码输出为聊天消息
        
        Args:
            output: 代码输出
            
        Returns:
            ChatMessage
        """
        content = f"**代码生成完成**\n\n{output.explanation}\n\n```python\n{output.code}\n```"
        
        if output.commands_to_run:
            content += "\n\n**需要执行的命令:**\n"
            for cmd in output.commands_to_run:
                content += f"- `{cmd}`\n"
        
        return ChatMessage(
            content=content,
            message_type="code",
            metadata={
                "code": output.code,
                "language": "python",
                "commands": output.commands_to_run
            }
        )

    def format_review(self, review: Review) -> ChatMessage:
        """
        格式化审查结果为聊天消息
        
        Args:
            review: 审查结果
            
        Returns:
            ChatMessage
        """
        status_emoji = "✅" if review.is_approved else "❌"
        reviewer_name = review.reviewer_type.capitalize()
        
        content = f"**{reviewer_name} Reviewer** {status_emoji}\n\n"
        content += f"- 状态: {review.status.value}\n"
        content += f"- 评分: {review.score}/10\n"
        
        if review.issues:
            content += "\n**问题:**\n"
            for issue in review.issues:
                content += f"- {issue}\n"
        
        if review.suggestions:
            content += "\n**建议:**\n"
            for suggestion in review.suggestions:
                content += f"- {suggestion}\n"
        
        return ChatMessage(
            content=content,
            message_type="review",
            metadata={
                "reviewer_type": review.reviewer_type,
                "status": review.status.value,
                "score": review.score
            }
        )

    def format_iteration_summary(self, result: IterationResult) -> ChatMessage:
        """
        格式化迭代摘要
        
        Args:
            result: 迭代结果
            
        Returns:
            ChatMessage
        """
        consensus_emoji = "🎉" if result.consensus_reached else "🔄"
        
        content = f"**迭代 {result.iteration_number} 完成** {consensus_emoji}\n\n"
        
        if result.consensus_reached:
            content += "所有 Reviewer 已通过！\n"
        else:
            content += "需要继续改进...\n"
        
        # 汇总评分
        if result.reviews:
            scores = [r.score for r in result.reviews]
            avg_score = sum(scores) / len(scores)
            content += f"\n平均评分: {avg_score:.1f}/10\n"
            
            # 各 Reviewer 状态
            content += "\n**Reviewer 状态:**\n"
            for review in result.reviews:
                status_icon = "✅" if review.is_approved else "❌"
                content += f"- {review.reviewer_type.capitalize()}: {status_icon} ({review.score}/10)\n"
        
        return ChatMessage(
            content=content,
            message_type="progress",
            metadata={
                "iteration": result.iteration_number,
                "consensus": result.consensus_reached
            }
        )


class AiideChatInterface:
    """与AIIDE侧边栏集成的接口"""

    def __init__(self):
        """初始化AIIDE聊天接口"""
        self.formatter = MessageFormatter()
        self._message_history: List[ChatMessage] = []

    def send_message(self, message: ChatMessage) -> None:
        """
        发送消息到侧边栏
        
        Args:
            message: 要发送的消息
        """
        self._message_history.append(message)
        # 实际实现中，这里会调用AIIDE的API发送消息
        self._output_to_sidebar(message)

    def send_text(self, text: str) -> None:
        """
        发送文本消息
        
        Args:
            text: 文本内容
        """
        message = ChatMessage(content=text, message_type="text")
        self.send_message(message)

    def send_code_block(self, code: str, language: str = "python") -> None:
        """
        发送带语法高亮的代码块
        
        Args:
            code: 代码内容
            language: 编程语言
        """
        content = f"```{language}\n{code}\n```"
        message = ChatMessage(
            content=content,
            message_type="code",
            metadata={"language": language, "code": code}
        )
        self.send_message(message)

    def send_review_feedback(self, review: Review) -> None:
        """
        发送格式化的审查反馈
        
        Args:
            review: 审查结果
        """
        message = self.formatter.format_review(review)
        self.send_message(message)

    def send_code_output(self, output: CodeOutput) -> None:
        """
        发送代码输出
        
        Args:
            output: 代码输出
        """
        message = self.formatter.format_code_output(output)
        self.send_message(message)

    def send_iteration_summary(self, result: IterationResult) -> None:
        """
        发送迭代摘要
        
        Args:
            result: 迭代结果
        """
        message = self.formatter.format_iteration_summary(result)
        self.send_message(message)

    def prompt_user(self, question: str, options: Optional[List[str]] = None) -> str:
        """
        提示用户输入并等待响应
        
        Args:
            question: 问题
            options: 可选的选项列表
            
        Returns:
            用户的响应
        """
        content = f"**{question}**"
        if options:
            content += "\n\n选项:\n"
            for i, opt in enumerate(options, 1):
                content += f"{i}. {opt}\n"
        
        message = ChatMessage(
            content=content,
            message_type="prompt",
            metadata={"options": options}
        )
        self.send_message(message)
        
        # 实际实现中，这里会等待用户输入
        # 模拟返回第一个选项或默认值
        return options[0] if options else "yes"

    def show_progress(self, iteration: int, phase: str) -> None:
        """
        显示当前进度
        
        Args:
            iteration: 当前迭代次数
            phase: 当前阶段
        """
        content = f"🔄 **迭代 {iteration}** - {phase}"
        message = ChatMessage(
            content=content,
            message_type="progress",
            metadata={"iteration": iteration, "phase": phase}
        )
        self.send_message(message)

    def _output_to_sidebar(self, message: ChatMessage) -> None:
        """
        输出到侧边栏（实际实现需要调用AIIDE API）
        
        Args:
            message: 消息
        """
        # 这里是模拟输出，实际实现需要调用AIIDE的侧边栏API
        print(message.content)

    def get_message_history(self) -> List[ChatMessage]:
        """获取消息历史"""
        return self._message_history.copy()

    def clear_history(self) -> None:
        """清空消息历史"""
        self._message_history.clear()
