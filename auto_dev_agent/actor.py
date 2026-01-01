"""
Actor Agent - 负责生成和修改代码的AI代理
"""

import json
import time
import re
from typing import List

from .models import Config, Task, CodeOutput, AggregatedFeedback


class ActorAgent:
    """负责生成和修改代码的代理"""

    def __init__(self, config: Config):
        """
        初始化Actor代理
        
        Args:
            config: 系统配置
        """
        self.config = config

    def generate_code(self, task: Task) -> CodeOutput:
        """
        根据任务描述生成初始代码
        
        Args:
            task: 任务描述
            
        Returns:
            CodeOutput 包含生成的代码
        """
        prompt = self._build_generation_prompt(task)
        
        for attempt in range(self.config.api_retry_count):
            try:
                response = self._call_ai_api(prompt)
                return self._parse_code_response(response)
            except Exception as e:
                if attempt < self.config.api_retry_count - 1:
                    time.sleep(2 ** attempt)
                else:
                    return CodeOutput(
                        code=f"# Error generating code: {str(e)}",
                        explanation=f"Failed to generate code after {self.config.api_retry_count} attempts",
                        commands_to_run=[]
                    )

    def revise_code(self, code: str, feedback: AggregatedFeedback, task: Task) -> CodeOutput:
        """
        根据反馈修改代码
        
        Args:
            code: 当前代码
            feedback: 聚合后的反馈
            task: 原始任务
            
        Returns:
            CodeOutput 包含修改后的代码
        """
        prompt = self._build_revision_prompt(code, feedback, task)
        
        for attempt in range(self.config.api_retry_count):
            try:
                response = self._call_ai_api(prompt)
                return self._parse_code_response(response)
            except Exception as e:
                if attempt < self.config.api_retry_count - 1:
                    time.sleep(2 ** attempt)
                else:
                    return CodeOutput(
                        code=code,  # Return original code on failure
                        explanation=f"Failed to revise code: {str(e)}",
                        commands_to_run=[]
                    )

    def _build_generation_prompt(self, task: Task) -> str:
        """构建代码生成提示词"""
        context_section = ""
        if task.context:
            context_section = f"\n现有代码/上下文：\n```\n{task.context}\n```\n"

        return f"""请根据以下任务描述生成代码。

任务描述：
{task.description}
{context_section}
请以JSON格式返回：
{{
    "code": "生成的代码",
    "explanation": "代码说明",
    "commands_to_run": ["需要执行的命令1", "命令2"]
}}

注意：
1. 代码应该完整可运行
2. 包含必要的注释
3. 遵循最佳实践
4. 如果需要安装依赖或运行测试，请在commands_to_run中列出
"""

    def _build_revision_prompt(self, code: str, feedback: AggregatedFeedback, task: Task) -> str:
        """构建代码修改提示词"""
        issues_text = "\n".join(f"- {issue}" for issue in feedback.combined_issues)
        suggestions_text = "\n".join(f"- {s}" for s in feedback.combined_suggestions)
        priority_text = "\n".join(f"- {p}" for p in feedback.priority_items)

        return f"""请根据审查反馈修改以下代码。

原始任务：
{task.description}

当前代码：
```
{code}
```

发现的问题：
{issues_text}

改进建议：
{suggestions_text}

优先处理项：
{priority_text}

请以JSON格式返回修改后的代码：
{{
    "code": "修改后的代码",
    "explanation": "修改说明",
    "commands_to_run": ["需要执行的命令"]
}}
"""

    def _call_ai_api(self, prompt: str) -> str:
        """
        调用AI API
        
        这是一个模拟实现，实际使用时需要替换为真实的API调用
        """
        # 模拟API调用 - 实际实现需要根据config.model_provider调用相应API
        return json.dumps({
            "code": "# Generated code placeholder\nprint('Hello, World!')",
            "explanation": "This is a placeholder implementation",
            "commands_to_run": []
        })

    def _parse_code_response(self, response: str) -> CodeOutput:
        """解析AI响应为CodeOutput对象"""
        try:
            data = json.loads(response)
            return CodeOutput(
                code=data.get("code", ""),
                explanation=data.get("explanation", ""),
                commands_to_run=data.get("commands_to_run", []),
                file_path=data.get("file_path")
            )
        except json.JSONDecodeError:
            # 尝试从响应中提取代码块
            code_match = re.search(r"```(?:\w+)?\n(.*?)```", response, re.DOTALL)
            if code_match:
                return CodeOutput(
                    code=code_match.group(1).strip(),
                    explanation="Extracted from response",
                    commands_to_run=[]
                )
            return CodeOutput(
                code=response,
                explanation="Raw response (failed to parse JSON)",
                commands_to_run=[]
            )
