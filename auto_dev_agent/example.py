"""
Auto Dev Agent 使用示例
"""

import sys
import os

# 支持直接运行
if __name__ == "__main__" and __package__ is None:
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from auto_dev_agent.models import Config, Task
    from auto_dev_agent.orchestrator import Orchestrator
    from auto_dev_agent.chat_interface import AiideChatInterface
else:
    from .models import Config, Task
    from .orchestrator import Orchestrator
    from .chat_interface import AiideChatInterface


def main():
    """示例：使用 Auto Dev Agent 完成一个简单任务"""
    
    # 1. 创建配置
    config = Config(
        model_provider="openai",
        model_name="gpt-4",
        max_iterations=5,
        safe_mode=True
    )
    
    # 2. 创建任务
    task = Task.create(
        description="实现一个Python函数，计算斐波那契数列的第n项",
        context=None
    )
    
    # 3. 初始化聊天接口
    chat = AiideChatInterface()
    
    # 4. 创建编排器
    orchestrator = Orchestrator(
        config=config,
        output_callback=chat.send_text,
        confirmation_callback=lambda cmd: True  # 自动确认命令
    )
    
    # 5. 执行任务
    print("开始执行任务...")
    print("=" * 50)
    
    result = orchestrator.run_task(task)
    
    # 6. 输出结果
    print("\n" + "=" * 50)
    print("任务完成!")
    print("=" * 50)
    print(f"\n成功: {result.success}")
    print(f"总迭代次数: {result.total_iterations}")
    print(f"\n摘要:\n{result.summary}")
    print(f"\n最终代码:\n{result.final_code}")


if __name__ == "__main__":
    main()
