"""
Kiro 自动运行器 - 固定提示词自动执行
"""

import time
import json
import hashlib
from dataclasses import dataclass, asdict
from typing import Optional, Tuple
from pathlib import Path

try:
    import pyautogui
    import pyperclip
    from PIL import ImageGrab
    HAS_DEPS = True
except ImportError:
    HAS_DEPS = False


@dataclass
class KiroConfig:
    """Kiro 自动化配置"""
    input_pos: Tuple[int, int] = (0, 0)       # 输入框位置
    run_button_pos: Tuple[int, int] = (0, 0)  # Run 按钮位置
    check_region: Tuple[int, int, int, int] = (0, 0, 100, 100)  # 检测区域 (x, y, w, h)
    stable_threshold: int = 50                  # 连续稳定次数判定完成（75 × 2秒 = 150秒）
    check_interval: float = 2.0                # 检测间隔（秒）
    max_wait: int = 300                        # 最大等待时间（秒）
    auto_click_run: bool = True                # 自动点击 Run 按钮
    run_check_interval: float = 20            # Run 按钮检测间隔
    
    def save(self, path: str = ".kiro_runner_config.json"):
        with open(path, "w") as f:
            json.dump(asdict(self), f, indent=2)
    
    @classmethod
    def load(cls, path: str = ".kiro_runner_config.json") -> "KiroConfig":
        try:
            with open(path, "r") as f:
                data = json.load(f)
                return cls(
                    input_pos=tuple(data["input_pos"]),
                    run_button_pos=tuple(data.get("run_button_pos", (0, 0))),
                    check_region=tuple(data.get("check_region", (0, 0, 100, 100))),
                    stable_threshold=data.get("stable_threshold", 3),
                    check_interval=data.get("check_interval", 2.0),
                    max_wait=data.get("max_wait", 300),
                    auto_click_run=data.get("auto_click_run", True),
                    run_check_interval=data.get("run_check_interval", 1.0)
                )
        except:
            return cls()


# ============ 固定提示词 ============
#FIXED_PROMPT = """我们现在要做市面上最好的Markdown编辑器，你作为一个专业的大厂程序员和一个使用者，请你批判性的提出建议，找出bug，找到不好看的地方，提出可以增加，优化的功能，并且自己测试，然后修改代码，请继续"""
FIXED_PROMPT = """请继续"""

class ScreenMonitor:
    """屏幕监控"""
    
    @staticmethod
    def capture_region(region: Tuple[int, int, int, int]) -> bytes:
        """截取区域"""
        x, y, w, h = region
        img = ImageGrab.grab(bbox=(x, y, x + w, y + h))
        return img.tobytes()
    
    @staticmethod
    def get_hash(data: bytes) -> str:
        """计算哈希"""
        return hashlib.md5(data).hexdigest()


class KiroRunner:
    """Kiro 自动运行器"""
    
    def __init__(self, config: Optional[KiroConfig] = None):
        if not HAS_DEPS:
            raise ImportError("请安装: pip install pyautogui pyperclip pillow")
        
        self.config = config or KiroConfig.load()
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.1
        self._run_clicked_count = 0
    
    def calibrate(self):
        """校准 - 记录输入框位置、Run按钮位置和检测区域"""
        print("=" * 50)
        print("Kiro Runner 校准")
        print("=" * 50)
        
        # 校准输入框
        print("\n[1/3] 请将鼠标移动到 Kiro 输入框，3秒后记录...")
        for i in range(3, 0, -1):
            print(f"  {i}...")
            time.sleep(1)
        pos = pyautogui.position()
        self.config.input_pos = (pos.x, pos.y)
        print(f"  ✓ 输入框位置: {pos}")
        
        # 校准 Run 按钮
        print("\n[2/3] 请将鼠标移动到 Run 按钮位置，3秒后记录...")
        print("      (LLM 询问是否执行命令时出现的 Run 按钮)")
        for i in range(3, 0, -1):
            print(f"  {i}...")
            time.sleep(1)
        pos = pyautogui.position()
        self.config.run_button_pos = (pos.x, pos.y)
        print(f"  ✓ Run按钮位置: {pos}")
        
        # 校准检测区域
        print("\n[3/3] 校准检测区域")
        print("      请将鼠标移动到聊天内容区域的左上角，3秒后记录...")
        for i in range(3, 0, -1):
            print(f"  {i}...")
            time.sleep(1)
        pos1 = pyautogui.position()
        
        print("\n      现在请将鼠标移动到聊天内容区域的右下角，3秒后记录...")
        for i in range(3, 0, -1):
            print(f"  {i}...")
            time.sleep(1)
        pos2 = pyautogui.position()
        
        self.config.check_region = (
            pos1.x, pos1.y,
            pos2.x - pos1.x, pos2.y - pos1.y
        )
        print(f"  ✓ 检测区域: {self.config.check_region}")
        
        # 保存
        self.config.save()
        print("\n✓ 配置已保存!")
        print("\n工作流程:")
        print("  1. 输入 prompt → 按 Enter 发送")
        print("  2. 等待过程中自动点击 Run 按钮（如果出现）")
        print("  3. 检测屏幕稳定后进入下一轮")
        return self.config
    
    def click_input(self):
        """点击输入框"""
        pyautogui.click(*self.config.input_pos)
        time.sleep(0.2)
    
    def clear_and_type(self, text: str):
        """清空并输入文本"""
        self.click_input()
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.1)
        pyperclip.copy(text)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.2)
    
    def send_enter(self):
        """按 Enter 发送消息"""
        pyautogui.press('enter')
        time.sleep(0.5)
    
    def click_run_button(self):
        """点击 Run 按钮"""
        pyautogui.click(*self.config.run_button_pos)
        self._run_clicked_count += 1
        time.sleep(0.3)
    
    def wait_and_auto_click_run(self) -> bool:
        """
        等待回答完成，期间自动点击 Run 按钮
        
        Returns:
            是否成功完成
        """
        print("等待回答完成（自动点击 Run）...")
        
        start_time = time.time()
        last_hash = None
        consecutive_same = 0
        self._run_clicked_count = 0
        
        while time.time() - start_time < self.config.max_wait:
            # 截图检测变化
            current_data = ScreenMonitor.capture_region(self.config.check_region)
            current_hash = ScreenMonitor.get_hash(current_data)
            
            if current_hash == last_hash:
                consecutive_same += 1
                print(f"\r  稳定 {consecutive_same}/{self.config.stable_threshold} | Run点击: {self._run_clicked_count}          ", end="", flush=True)
                
                if consecutive_same >= self.config.stable_threshold:
                    print()
                    return True
            else:
                consecutive_same = 0
                print(f"\r  内容变化中... | Run点击: {self._run_clicked_count}          ", end="", flush=True)
                
                # 内容变化时，尝试点击 Run 按钮
                if self.config.auto_click_run:
                    self.click_run_button()
            
            last_hash = current_hash
            time.sleep(self.config.check_interval)
        
        print()
        return False
    
    def execute_once(self, prompt: str = None) -> bool:
        """
        执行一次完整流程：
        1. 输入 prompt
        2. 按 Enter 发送
        3. 等待完成（期间自动点击 Run）
        """
        prompt = prompt or FIXED_PROMPT
        
        # 输入
        print(f"输入: {prompt[:40]}...")
        self.clear_and_type(prompt)
        
        # 按 Enter 发送
        print("按 Enter 发送...")
        self.send_enter()
        
        # 等待完成（自动点击 Run）
        success = self.wait_and_auto_click_run()
        
        if success:
            print(f"✓ 完成 (Run 点击了 {self._run_clicked_count} 次)")
        else:
            print("⚠ 等待超时")
        
        return success
    
    def run_loop(self, iterations: int = 10, prompt: str = None):
        """循环执行"""
        prompt = prompt or FIXED_PROMPT
        
        print(f"\n开始循环执行，共 {iterations} 次")
        print(f"提示词: {prompt[:50]}...")
        print("流程: 输入 → Enter发送 → 自动点Run → 等待完成")
        print("按 Ctrl+C 或将鼠标移到屏幕左上角可中断")
        print("=" * 50)
        
        total_run_clicks = 0
        
        for i in range(1, iterations + 1):
            try:
                print(f"\n[{i}/{iterations}] 执行中...")
                success = self.execute_once(prompt)
                total_run_clicks += self._run_clicked_count
                
                if not success:
                    print("⚠ 本次可能未完成，继续...")
                
                time.sleep(1)
                    
            except pyautogui.FailSafeException:
                print("\n⚠ 检测到鼠标在左上角，已中断")
                break
            except KeyboardInterrupt:
                print("\n⚠ 用户中断")
                break
        
        print("\n" + "=" * 50)
        print(f"✓ 完成，共点击 Run {total_run_clicks} 次")


def launch_in_new_window():
    """在新窗口中启动（不占用当前终端）"""
    import subprocess
    import sys
    import os
    
    # 获取当前脚本路径
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 构建命令
    args = sys.argv[2:] if len(sys.argv) > 2 else ["10"]
    cmd = f'python -m auto_dev_agent.kiro_runner _run {" ".join(args)}'
    
    # 在新的 cmd 窗口中运行
    subprocess.Popen(
        f'start cmd /k "{cmd}"',
        shell=True,
        cwd=os.path.dirname(script_dir)
    )
    print("✓ 已在新窗口中启动 Kiro Runner")
    print("  可以关闭此终端，新窗口会继续运行")


def main():
    """主函数"""
    import sys
    
    if not HAS_DEPS:
        print("请先安装依赖: pip install pyautogui pyperclip pillow")
        return
    
    runner = KiroRunner()
    
    if len(sys.argv) > 1:
        cmd = sys.argv[1]
        
        if cmd == "calibrate":
            runner.calibrate()
            
        elif cmd == "run":
            # 在新窗口中启动，不占用当前终端
            launch_in_new_window()
            
        elif cmd == "_run":
            # 实际运行（由新窗口调用）
            if runner.config.input_pos == (0, 0):
                print("请先校准: python -m auto_dev_agent.kiro_runner calibrate")
                return
            
            iterations = int(sys.argv[2]) if len(sys.argv) > 2 else 10
            
            print("3秒后开始...")
            time.sleep(3)
            runner.run_loop(iterations=iterations)
            
        elif cmd == "once":
            if runner.config.input_pos == (0, 0):
                print("请先校准: python -m auto_dev_agent.kiro_runner calibrate")
                return
            print("3秒后执行...")
            time.sleep(3)
            runner.execute_once()
            
        elif cmd == "status":
            print("当前配置:")
            print(f"  输入框位置: {runner.config.input_pos}")
            print(f"  Run按钮位置: {runner.config.run_button_pos}")
            print(f"  检测区域: {runner.config.check_region}")
            print(f"  自动点击Run: {runner.config.auto_click_run}")
        
        elif cmd == "stop":
            print("要停止运行，请在 Kiro Runner 窗口中按 Ctrl+C")
            print("或将鼠标移到屏幕左上角")
            
        else:
            print_usage()
    else:
        print_usage()


def print_usage():
    print("""
Kiro 自动运行器

用法:
  python -m auto_dev_agent.kiro_runner calibrate    # 校准位置
  python -m auto_dev_agent.kiro_runner run [次数]   # 在新窗口中运行（不占用终端）
  python -m auto_dev_agent.kiro_runner once         # 执行一次（占用终端）
  python -m auto_dev_agent.kiro_runner status       # 查看当前配置

工作流程:
  1. 输入 prompt → 按 Enter 发送
  2. 等待过程中自动点击 Run 按钮（LLM询问执行命令时）
  3. 屏幕稳定后进入下一轮

示例:
  python -m auto_dev_agent.kiro_runner calibrate
  python -m auto_dev_agent.kiro_runner run 10       # 在新窗口运行10次

停止运行:
  在 Kiro Runner 窗口按 Ctrl+C，或将鼠标移到屏幕左上角

修改提示词:
  编辑 auto_dev_agent/kiro_runner.py 中的 FIXED_PROMPT 变量
""")


if __name__ == "__main__":
    main()
