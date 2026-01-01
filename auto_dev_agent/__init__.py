"""
Auto Dev Agent - Actor-Reviewer 自动化开发代理系统

一个使用 Actor-Reviewer 架构的自动化 AI 开发代理，
通过多轮迭代自动完成代码开发任务。
"""

__version__ = "0.1.0"

# Lazy imports to avoid circular dependencies during testing
def __getattr__(name):
    if name == "Config":
        from .models import Config
        return Config
    elif name == "Task":
        from .models import Task
        return Task
    elif name == "CodeOutput":
        from .models import CodeOutput
        return CodeOutput
    elif name == "Review":
        from .models import Review
        return Review
    elif name == "ReviewStatus":
        from .models import ReviewStatus
        return ReviewStatus
    elif name == "AggregatedFeedback":
        from .models import AggregatedFeedback
        return AggregatedFeedback
    elif name == "IterationResult":
        from .models import IterationResult
        return IterationResult
    elif name == "DevState":
        from .models import DevState
        return DevState
    elif name == "TaskResult":
        from .models import TaskResult
        return TaskResult
    elif name == "CommandResult":
        from .models import CommandResult
        return CommandResult
    elif name == "ChatMessage":
        from .models import ChatMessage
        return ChatMessage
    elif name == "Orchestrator":
        from .orchestrator import Orchestrator
        return Orchestrator
    elif name == "ActorAgent":
        from .actor import ActorAgent
        return ActorAgent
    elif name == "CorrectnessReviewer":
        from .reviewers import CorrectnessReviewer
        return CorrectnessReviewer
    elif name == "StyleReviewer":
        from .reviewers import StyleReviewer
        return StyleReviewer
    elif name == "RobustnessReviewer":
        from .reviewers import RobustnessReviewer
        return RobustnessReviewer
    elif name == "StateManager":
        from .state_manager import StateManager
        return StateManager
    elif name == "CommandExecutor":
        from .command_executor import CommandExecutor
        return CommandExecutor
    elif name == "AiideChatInterface":
        from .chat_interface import AiideChatInterface
        return AiideChatInterface
    elif name == "MessageFormatter":
        from .chat_interface import MessageFormatter
        return MessageFormatter
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")

__all__ = [
    "Config",
    "Task",
    "CodeOutput",
    "Review",
    "ReviewStatus",
    "AggregatedFeedback",
    "IterationResult",
    "DevState",
    "TaskResult",
    "CommandResult",
    "ChatMessage",
    "Orchestrator",
    "ActorAgent",
    "CorrectnessReviewer",
    "StyleReviewer",
    "RobustnessReviewer",
    "StateManager",
    "CommandExecutor",
    "AiideChatInterface",
    "MessageFormatter",
]
