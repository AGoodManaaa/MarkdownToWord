"""
命令执行器 - 负责执行shell命令
"""

import subprocess
import re
import locale
from typing import List, Callable, Optional

from .models import CommandResult


class CommandExecutor:
    """执行shell命令的组件"""

    # 危险命令模式
    DANGEROUS_PATTERNS = [
        r"rm\s+-rf\s+/",
        r"rm\s+-rf\s+\*",
        r"rm\s+-rf\s+~",
        r"mkfs\.",
        r"dd\s+if=",
        r":\(\)\{:\|:&\};:",  # Fork bomb
        r">\s*/dev/sd",
        r"chmod\s+-R\s+777\s+/",
        r"chown\s+-R.*\s+/",
        r"curl.*\|\s*sh",
        r"wget.*\|\s*sh",
        r"format\s+[a-z]:",  # Windows format
        r"del\s+/[fqs]\s+",  # Windows delete
    ]

    def __init__(
        self,
        safe_mode: bool = True,
        timeout: int = 60,
        confirmation_callback: Optional[Callable[[str], bool]] = None
    ):
        """
        初始化命令执行器
        
        Args:
            safe_mode: 是否启用安全模式（危险命令需确认）
            timeout: 命令执行超时时间（秒）
            confirmation_callback: 危险命令确认回调函数
        """
        self.safe_mode = safe_mode
        self.timeout = timeout
        self.confirmation_callback = confirmation_callback
        self._dangerous_patterns = [re.compile(p, re.IGNORECASE) for p in self.DANGEROUS_PATTERNS]

    def is_dangerous(self, command: str) -> bool:
        """
        检查命令是否危险
        
        Args:
            command: 要检查的命令
            
        Returns:
            是否为危险命令
        """
        for pattern in self._dangerous_patterns:
            if pattern.search(command):
                return True
        return False

    def execute(self, command: str) -> CommandResult:
        """
        执行shell命令
        
        Args:
            command: 要执行的命令
            
        Returns:
            CommandResult 包含执行结果
        """
        # 检查危险命令
        if self.safe_mode and self.is_dangerous(command):
            if self.confirmation_callback:
                confirmed = self.confirmation_callback(command)
                if not confirmed:
                    return CommandResult(
                        command=command,
                        success=False,
                        stdout="",
                        stderr="Command execution cancelled: dangerous command requires confirmation",
                        return_code=-1
                    )
            else:
                return CommandResult(
                    command=command,
                    success=False,
                    stdout="",
                    stderr="Command execution blocked: dangerous command detected and no confirmation callback provided",
                    return_code=-1
                )

        try:
            preferred_encoding = locale.getpreferredencoding(False) or "utf-8"
            result = subprocess.run(
                command,
                shell=True,
                capture_output=True,
                text=True,
                encoding=preferred_encoding,
                errors="replace",
                timeout=self.timeout
            )
            
            return CommandResult(
                command=command,
                success=result.returncode == 0,
                stdout=result.stdout or "",
                stderr=result.stderr or "",
                return_code=result.returncode
            )
        except subprocess.TimeoutExpired:
            return CommandResult(
                command=command,
                success=False,
                stdout="",
                stderr=f"Command timed out after {self.timeout} seconds",
                return_code=-2
            )
        except Exception as e:
            return CommandResult(
                command=command,
                success=False,
                stdout="",
                stderr=str(e),
                return_code=-3
            )

    def execute_multiple(self, commands: List[str]) -> List[CommandResult]:
        """
        执行多个命令
        
        Args:
            commands: 命令列表
            
        Returns:
            CommandResult 列表
        """
        results = []
        for cmd in commands:
            result = self.execute(cmd)
            results.append(result)
            # 如果命令失败，继续执行但记录错误
        return results
