"""
Reviewer Agents - 负责审查代码的AI代理
"""

import json
import time
from abc import ABC, abstractmethod
from typing import Optional

from .models import Config, Task, Review, ReviewStatus


class BaseReviewer(ABC):
    """Reviewer基类"""

    def __init__(self, config: Config, focus_area: str):
        """
        初始化Reviewer
        
        Args:
            config: 系统配置
            focus_area: 审查关注领域
        """
        self.config = config
        self.focus_area = focus_area
        self.reviewer_type = self.__class__.__name__.replace("Reviewer", "").lower()

    @abstractmethod
    def get_system_prompt(self) -> str:
        """获取系统提示词"""
        pass

    def review(self, code: str, task: Task) -> Review:
        """
        审查代码并返回反馈
        
        Args:
            code: 要审查的代码
            task: 任务描述
            
        Returns:
            Review 审查结果
        """
        prompt = self._build_review_prompt(code, task)
        
        for attempt in range(self.config.api_retry_count):
            try:
                response = self._call_ai_api(prompt)
                return self._parse_review_response(response)
            except Exception as e:
                if attempt < self.config.api_retry_count - 1:
                    time.sleep(2 ** attempt)  # Exponential backoff
                else:
                    # Return a default review on failure
                    return Review(
                        reviewer_type=self.reviewer_type,
                        status=ReviewStatus.NEEDS_REVISION,
                        issues=[f"AI API error: {str(e)}"],
                        suggestions=["Please retry the review"],
                        score=1
                    )

    def _build_review_prompt(self, code: str, task: Task) -> str:
        """构建审查提示词"""
        return f"""请审查以下代码，关注{self.focus_area}。

任务描述：
{task.description}

代码：
```
{code}
```

请以JSON格式返回审查结果：
{{
    "status": "approved" 或 "needs_revision",
    "issues": ["问题1", "问题2"],
    "suggestions": ["建议1", "建议2"],
    "score": 1-10的评分
}}
"""

    def _call_ai_api(self, prompt: str) -> str:
        """
        调用AI API
        
        这是一个模拟实现，实际使用时需要替换为真实的API调用
        """
        # 模拟API调用 - 实际实现需要根据config.model_provider调用相应API
        # 这里返回一个默认的审查通过响应
        return json.dumps({
            "status": "approved",
            "issues": [],
            "suggestions": [],
            "score": 8
        })

    def _parse_review_response(self, response: str) -> Review:
        """解析AI响应为Review对象"""
        try:
            data = json.loads(response)
            status = ReviewStatus.APPROVED if data.get("status") == "approved" else ReviewStatus.NEEDS_REVISION
            
            issues = data.get("issues", [])
            suggestions = data.get("suggestions", [])
            
            # 确保有问题时必须有建议
            if issues and not suggestions:
                suggestions = ["Please address the identified issues"]
            
            return Review(
                reviewer_type=self.reviewer_type,
                status=status,
                issues=issues,
                suggestions=suggestions,
                score=data.get("score", 5)
            )
        except json.JSONDecodeError:
            return Review(
                reviewer_type=self.reviewer_type,
                status=ReviewStatus.NEEDS_REVISION,
                issues=["Failed to parse AI response"],
                suggestions=["Please retry the review"],
                score=1
            )


class CorrectnessReviewer(BaseReviewer):
    """专注于代码正确性和逻辑错误的Reviewer"""

    def __init__(self, config: Config):
        super().__init__(config, "代码正确性和逻辑错误")

    def get_system_prompt(self) -> str:
        return """你是一个专注于代码正确性的审查专家。
你的职责是：
1. 检查代码逻辑是否正确
2. 发现潜在的bug和错误
3. 验证算法实现是否符合预期
4. 检查边界条件处理
5. 确保代码能正确完成任务要求"""


class StyleReviewer(BaseReviewer):
    """专注于代码风格和最佳实践的Reviewer"""

    def __init__(self, config: Config):
        super().__init__(config, "代码风格、可读性和最佳实践")

    def get_system_prompt(self) -> str:
        return """你是一个专注于代码风格的审查专家。
你的职责是：
1. 检查代码是否遵循编码规范
2. 评估代码可读性和可维护性
3. 检查命名是否清晰合理
4. 评估代码结构和组织
5. 建议最佳实践改进"""


class RobustnessReviewer(BaseReviewer):
    """专注于边界情况和错误处理的Reviewer"""

    def __init__(self, config: Config):
        super().__init__(config, "边界情况、错误处理和健壮性")

    def get_system_prompt(self) -> str:
        return """你是一个专注于代码健壮性的审查专家。
你的职责是：
1. 检查边界情况处理
2. 评估错误处理机制
3. 检查异常情况的处理
4. 评估代码的防御性编程
5. 检查资源管理和清理"""
