"""
Auto Dev Agent - 主入口点
"""

import argparse
import sys
import os
from typing import Optional

# 支持直接运行和作为模块运行
if __name__ == "__main__" and __package__ is None:
    # 直接运行时，添加父目录到路径
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from auto_dev_agent.models import Config, Task
    from auto_dev_agent.orchestrator import Orchestrator
    from auto_dev_agent.state_manager import StateManager
    from auto_dev_agent.chat_interface import AiideChatInterface
else:
    from .models import Config, Task
    from .orchestrator import Orchestrator
    from .state_manager import StateManager
    from .chat_interface import AiideChatInterface


def create_parser() -> argparse.ArgumentParser:
    """创建命令行参数解析器"""
    parser = argparse.ArgumentParser(
        description="Auto Dev Agent - Actor-Reviewer 自动化开发代理",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  %(prog)s "实现一个快速排序算法"
  %(prog)s "创建一个REST API端点" --max-iterations 5
  %(prog)s --resume task_id
  %(prog)s --list-pending
        """
    )
    
    parser.add_argument(
        "task",
        nargs="?",
        help="任务描述"
    )
    
    parser.add_argument(
        "--context", "-c",
        help="额外上下文（如现有代码文件路径）"
    )
    
    parser.add_argument(
        "--max-iterations", "-m",
        type=int,
        default=10,
        help="最大迭代次数（默认: 10）"
    )
    
    parser.add_argument(
        "--model-provider",
        default="openai",
        choices=["openai", "anthropic", "azure", "local"],
        help="AI模型提供商（默认: openai）"
    )
    
    parser.add_argument(
        "--model-name",
        default="gpt-4",
        help="AI模型名称（默认: gpt-4）"
    )
    
    parser.add_argument(
        "--unsafe",
        action="store_true",
        help="禁用安全模式（允许执行危险命令）"
    )
    
    parser.add_argument(
        "--state-dir",
        default=".auto_dev_state",
        help="状态文件目录（默认: .auto_dev_state）"
    )
    
    parser.add_argument(
        "--resume", "-r",
        metavar="TASK_ID",
        help="恢复指定任务"
    )
    
    parser.add_argument(
        "--list-pending", "-l",
        action="store_true",
        help="列出所有未完成的任务"
    )
    
    return parser


def load_context(context_path: Optional[str]) -> Optional[str]:
    """加载上下文文件"""
    if not context_path:
        return None
    
    try:
        with open(context_path, "r", encoding="utf-8") as f:
            return f.read()
    except FileNotFoundError:
        print(f"警告: 上下文文件不存在: {context_path}")
        return None
    except Exception as e:
        print(f"警告: 读取上下文文件失败: {e}")
        return None


def confirmation_callback(command: str) -> bool:
    """危险命令确认回调"""
    print(f"\n⚠️  检测到危险命令: {command}")
    response = input("是否继续执行? (yes/no): ").strip().lower()
    return response in ("yes", "y")


def main():
    """主函数"""
    parser = create_parser()
    args = parser.parse_args()
    
    # 创建配置
    config = Config(
        model_provider=args.model_provider,
        model_name=args.model_name,
        max_iterations=args.max_iterations,
        safe_mode=not args.unsafe,
        state_dir=args.state_dir
    )
    
    # 初始化状态管理器
    state_manager = StateManager(config.state_dir)
    
    # 列出未完成任务
    if args.list_pending:
        pending = state_manager.list_pending_tasks()
        if pending:
            print("未完成的任务:")
            for task_id in pending:
                print(f"  - {task_id}")
        else:
            print("没有未完成的任务")
        return 0
    
    # 初始化聊天接口
    chat = AiideChatInterface()
    
    # 创建编排器
    orchestrator = Orchestrator(
        config=config,
        output_callback=chat.send_text,
        confirmation_callback=confirmation_callback
    )
    
    # 恢复任务
    if args.resume:
        state = state_manager.load_state(args.resume)
        if state:
            print(f"恢复任务: {args.resume}")
            result = orchestrator.run_task(state.task, resume_state=state)
        else:
            print(f"错误: 找不到任务 {args.resume}")
            return 1
    else:
        # 新任务
        if not args.task:
            parser.print_help()
            return 1
        
        context = load_context(args.context)
        task = Task.create(description=args.task, context=context)
        
        # 检查是否有未完成的任务
        pending = state_manager.list_pending_tasks()
        if pending:
            print(f"发现 {len(pending)} 个未完成的任务")
            response = input("是否继续新任务? (yes/no): ").strip().lower()
            if response not in ("yes", "y"):
                print("使用 --resume TASK_ID 恢复任务，或 --list-pending 查看列表")
                return 0
        
        result = orchestrator.run_task(task)
    
    # 输出结果
    print("\n" + "=" * 50)
    print("任务完成!")
    print("=" * 50)
    print(result.summary)
    
    if result.success:
        print("\n最终代码:")
        print("-" * 50)
        print(result.final_code)
        return 0
    else:
        print("\n任务未能完全完成，请查看上述输出了解详情")
        return 1


if __name__ == "__main__":
    sys.exit(main())
