"""
数据模型定义
"""

from dataclasses import dataclass, field, asdict
from typing import List, Optional, Dict, Any
from enum import Enum
import json
import uuid
from datetime import datetime


class ReviewStatus(Enum):
    """审查状态枚举"""
    APPROVED = "approved"
    NEEDS_REVISION = "needs_revision"


@dataclass
class Config:
    """系统配置"""
    model_provider: str = "openai"
    model_name: str = "gpt-4"
    max_iterations: int = 10
    safe_mode: bool = True
    state_dir: str = ".auto_dev_state"
    command_timeout: int = 60
    api_retry_count: int = 3

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "Config":
        return cls(**data)

    @classmethod
    def get_default(cls) -> "Config":
        return cls()


@dataclass
class Task:
    """开发任务"""
    id: str
    description: str
    context: Optional[str] = None
    created_at: str = field(default_factory=lambda: datetime.now().isoformat())

    @classmethod
    def create(cls, description: str, context: Optional[str] = None) -> "Task":
        return cls(
            id=str(uuid.uuid4())[:8],
            description=description,
            context=context
        )

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "Task":
        return cls(**data)


@dataclass
class CodeOutput:
    """Actor生成的代码输出"""
    code: str
    explanation: str
    commands_to_run: List[str] = field(default_factory=list)
    file_path: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "CodeOutput":
        return cls(**data)


@dataclass
class Review:
    """Reviewer的审查结果"""
    reviewer_type: str
    status: ReviewStatus
    issues: List[str] = field(default_factory=list)
    suggestions: List[str] = field(default_factory=list)
    score: int = 5

    def to_dict(self) -> Dict[str, Any]:
        d = asdict(self)
        d["status"] = self.status.value
        return d

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "Review":
        data = data.copy()
        data["status"] = ReviewStatus(data["status"])
        return cls(**data)

    @property
    def is_approved(self) -> bool:
        return self.status == ReviewStatus.APPROVED


@dataclass
class AggregatedFeedback:
    """聚合后的反馈"""
    all_approved: bool
    combined_issues: List[str] = field(default_factory=list)
    combined_suggestions: List[str] = field(default_factory=list)
    priority_items: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "AggregatedFeedback":
        return cls(**data)


@dataclass
class IterationResult:
    """单次迭代的结果"""
    iteration_number: int
    code: str
    reviews: List[Review] = field(default_factory=list)
    consensus_reached: bool = False
    command_results: List["CommandResult"] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "iteration_number": self.iteration_number,
            "code": self.code,
            "reviews": [r.to_dict() for r in self.reviews],
            "consensus_reached": self.consensus_reached,
            "command_results": [c.to_dict() for c in self.command_results],
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "IterationResult":
        return cls(
            iteration_number=data["iteration_number"],
            code=data["code"],
            reviews=[Review.from_dict(r) for r in data.get("reviews", [])],
            consensus_reached=data.get("consensus_reached", False),
            command_results=[CommandResult.from_dict(c) for c in data.get("command_results", [])],
        )


@dataclass
class DevState:
    """开发状态"""
    task: Task
    current_iteration: int
    current_code: str
    iteration_history: List[IterationResult] = field(default_factory=list)
    status: str = "in_progress"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "task": self.task.to_dict(),
            "current_iteration": self.current_iteration,
            "current_code": self.current_code,
            "iteration_history": [i.to_dict() for i in self.iteration_history],
            "status": self.status,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "DevState":
        return cls(
            task=Task.from_dict(data["task"]),
            current_iteration=data["current_iteration"],
            current_code=data["current_code"],
            iteration_history=[IterationResult.from_dict(i) for i in data.get("iteration_history", [])],
            status=data.get("status", "in_progress"),
        )

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=2)

    @classmethod
    def from_json(cls, json_str: str) -> "DevState":
        return cls.from_dict(json.loads(json_str))


@dataclass
class TaskResult:
    """任务最终结果"""
    success: bool
    final_code: str
    total_iterations: int
    final_reviews: List[Review] = field(default_factory=list)
    summary: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "success": self.success,
            "final_code": self.final_code,
            "total_iterations": self.total_iterations,
            "final_reviews": [r.to_dict() for r in self.final_reviews],
            "summary": self.summary,
        }


@dataclass
class CommandResult:
    """命令执行结果"""
    command: str
    success: bool
    stdout: str = ""
    stderr: str = ""
    return_code: int = 0

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "CommandResult":
        return cls(**data)


@dataclass
class ChatMessage:
    """聊天消息"""
    content: str
    message_type: str = "text"
    metadata: Optional[Dict[str, Any]] = None

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)
