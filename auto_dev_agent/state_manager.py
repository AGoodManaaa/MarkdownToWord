"""
状态管理器 - 负责任务状态的持久化
"""

import os
import json
from pathlib import Path
from typing import Optional, List

from .models import DevState, Task


class StateManager:
    """管理任务状态的持久化"""

    def __init__(self, state_dir: str = ".auto_dev_state"):
        """
        初始化状态管理器
        
        Args:
            state_dir: 状态文件存储目录
        """
        self.state_dir = Path(state_dir)
        self._ensure_dir_exists()

    def _ensure_dir_exists(self) -> None:
        """确保状态目录存在"""
        self.state_dir.mkdir(parents=True, exist_ok=True)

    def _get_state_file_path(self, task_id: str) -> Path:
        """获取任务状态文件路径"""
        return self.state_dir / f"{task_id}.json"

    def save_state(self, state: DevState) -> None:
        """
        保存当前状态到JSON文件
        
        Args:
            state: 要保存的开发状态
        """
        file_path = self._get_state_file_path(state.task.id)
        json_str = state.to_json()
        
        # 先写入临时文件，再重命名，确保原子性
        temp_path = file_path.with_suffix(".tmp")
        try:
            with open(temp_path, "w", encoding="utf-8") as f:
                f.write(json_str)
            temp_path.replace(file_path)
        except Exception as e:
            if temp_path.exists():
                temp_path.unlink()
            raise RuntimeError(f"Failed to save state: {e}")

    def load_state(self, task_id: str) -> Optional[DevState]:
        """
        加载已保存的状态
        
        Args:
            task_id: 任务ID
            
        Returns:
            DevState if found, None otherwise
        """
        file_path = self._get_state_file_path(task_id)
        
        if not file_path.exists():
            return None
        
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                json_str = f.read()
            return DevState.from_json(json_str)
        except (json.JSONDecodeError, KeyError) as e:
            raise RuntimeError(f"Failed to load state (corrupted file): {e}")

    def list_pending_tasks(self) -> List[str]:
        """
        列出所有未完成的任务
        
        Returns:
            未完成任务的ID列表
        """
        pending = []
        
        if not self.state_dir.exists():
            return pending
        
        for file_path in self.state_dir.glob("*.json"):
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if data.get("status") == "in_progress":
                    pending.append(data["task"]["id"])
            except (json.JSONDecodeError, KeyError):
                continue
        
        return pending

    def delete_state(self, task_id: str) -> bool:
        """
        删除任务状态文件
        
        Args:
            task_id: 任务ID
            
        Returns:
            是否成功删除
        """
        file_path = self._get_state_file_path(task_id)
        
        if file_path.exists():
            file_path.unlink()
            return True
        return False

    def state_exists(self, task_id: str) -> bool:
        """
        检查任务状态是否存在
        
        Args:
            task_id: 任务ID
            
        Returns:
            状态文件是否存在
        """
        return self._get_state_file_path(task_id).exists()

    def get_all_task_ids(self) -> List[str]:
        """
        获取所有任务ID
        
        Returns:
            所有任务ID列表
        """
        task_ids = []
        
        if not self.state_dir.exists():
            return task_ids
        
        for file_path in self.state_dir.glob("*.json"):
            task_ids.append(file_path.stem)
        
        return task_ids
