"""
Kiro Visual Automation - 通过视觉自动化与Kiro侧边栏交互
"""

import time
import json
from dataclasses import dataclass, asdict
from typing import Optional, Tuple, List
from pathlib import Path

try:
    import pyautogui
    import pyperclip
    from PIL import Image, ImageGrab
    HAS_VISUAL_DEPS = True
except ImportError:
    HAS_VISUAL_DEPS = False


@dataclass
class VisualConfig:
    """视觉自动化配置"""
    input_field_pos: Tuple[int, int] = (0, 0)  # 输入框点击位置
    typing_interval: float = 0.01  # 打字间隔
    response_timeout: int = 120  # 响应超时（秒）
    check_interval: float = 1.0  # 检查间隔
    use_clipboard: bool = True  # 使用剪贴板粘贴（更快）
    
    def to_dict(self):
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: dict) -> "VisualConfig":
        return cls(**data)
    
    def save(self, path: str = ".kiro_visual_config.json"):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, indent=2)
    
    @classmethod
    def load(cls, path: str = ".kiro_visual_config.json") -> "VisualConfig":
        try:
            with open(path, "r", encoding="utf-8") as f:
                return cls.from_dict(json.load(f))
        except FileNotFoundError:
            return cls()


class ScreenCapture:
    """屏幕截图和图像处理"""
    
    def __init__(self):
        if not HAS_VISUAL_DEPS:
            raise ImportError("需要安装: pip install pyautogui pillow pyperclip")
    
    def capture_full(self) -> Image.Image:
        """截取全屏"""
        return ImageGrab.grab()
    
    def capture_region(self, x: int, y: int, width: int, height: int) -> Image.Image:
        """截取指定区域"""
        return ImageGrab.grab(bbox=(x, y, x + width, y + height))
    
    def get_screen_size(self) -> Tuple[int, int]:
        """获取屏幕尺寸"""
        return pyautogui.size()


class InputSimulator:
    """键盘鼠标模拟"""
    
    def __init__(self):
        if not HAS_VISUAL_DEPS:
            raise ImportError("需要安装: pip install pyautogui pillow pyperclip")
        # 设置安全特性
        pyautogui.FAILSAFE = True  # 移动到左上角可中断
        pyautogui.PAUSE = 0.05  # 操作间隔
    
    def click(self, x: int, y: int, clicks: int = 1) -> None:
        """模拟鼠标点击"""
        pyautogui.click(x, y, clicks=clicks)
    
    def type_text(self, text: str, interval: float = 0.02) -> None:
        """模拟键盘输入（逐字符）"""
        pyautogui.typewrite(text, interval=interval)
    
    def type_text_unicode(self, text: str, interval: float = 0.01) -> None:
        """输入Unicode文本（支持中文）"""
        for char in text:
            pyautogui.write(char)
            time.sleep(interval)
    
    def paste_text(self, text: str) -> None:
        """通过剪贴板粘贴文本（更快，支持中文）"""
        pyperclip.copy(text)
        pyautogui.hotkey('ctrl', 'v')
    
    def press_key(self, key: str) -> None:
        """模拟按键"""
        pyautogui.press(key)
    
    def hotkey(self, *keys) -> None:
        """模拟组合键"""
        pyautogui.hotkey(*keys)
    
    def select_all(self) -> None:
        """全选"""
        pyautogui.hotkey('ctrl', 'a')
    
    def clear_input(self) -> None:
        """清空输入框"""
        self.select_all()
        self.press_key('delete')


class KiroVisualInterface:
    """通过视觉自动化与Kiro侧边栏交互"""
    
    def __init__(self, config: Optional[VisualConfig] = None):
        """
        初始化视觉自动化接口
        
        Args:
            config: 视觉配置，如果为None则尝试加载或使用默认值
        """
        if not HAS_VISUAL_DEPS:
            raise ImportError("需要安装: pip install pyautogui pillow pyperclip")
        
        self.config = config or VisualConfig.load()
        self.screen = ScreenCapture()
        self.input = InputSimulator()
        self._last_response = ""
    
    def calibrate(self) -> VisualConfig:
        """
        交互式校准 - 让用户点击输入框位置
        
        Returns:
            校准后的配置
        """
        print("=" * 50)
        print("Kiro 视觉自动化校准")
        print("=" * 50)
        print("\n请在 3 秒内将鼠标移动到 Kiro 输入框位置...")
        
        for i in range(3, 0, -1):
            print(f"  {i}...")
            time.sleep(1)
        
        pos = pyautogui.position()
        self.config.input_field_pos = (pos.x, pos.y)
        
        print(f"\n✓ 已记录输入框位置: {pos}")
        
        # 保存配置
        self.config.save()
        print(f"✓ 配置已保存到 .kiro_visual_config.json")
        
        return self.config
    
    def focus_input(self) -> None:
        """点击输入框获取焦点"""
        x, y = self.config.input_field_pos
        self.input.click(x, y)
        time.sleep(0.1)
    
    def type_message(self, message: str) -> None:
        """
        在输入框中输入消息
        
        Args:
            message: 要输入的消息
        """
        self.focus_input()
        self.input.clear_input()
        time.sleep(0.1)
        
        if self.config.use_clipboard:
            # 使用剪贴板粘贴（更快，支持中文）
            self.input.paste_text(message)
        else:
            # 逐字符输入
            self.input.type_text_unicode(message, self.config.typing_interval)
    
    def submit_message(self) -> None:
        """提交消息（按Enter）"""
        time.sleep(0.1)
        self.input.press_key('enter')
    
    def send_message(self, message: str) -> None:
        """
        发送消息到Kiro（输入+提交）
        
        Args:
            message: 要发送的消息
        """
        self.type_message(message)
        time.sleep(0.2)
        self.submit_message()
    
    def wait_for_response(self, timeout: Optional[int] = None) -> bool:
        """
        等待Kiro响应完成
        
        简单实现：等待固定时间
        高级实现可以通过屏幕变化检测来判断响应是否完成
        
        Args:
            timeout: 超时时间（秒）
            
        Returns:
            是否在超时前完成
        """
        timeout = timeout or self.config.response_timeout
        print(f"等待 Kiro 响应（最多 {timeout} 秒）...")
        
        # 简单实现：等待固定时间
        # TODO: 可以通过检测屏幕变化来判断响应是否完成
        time.sleep(timeout)
        return True
    
    def send_and_wait(self, message: str, wait_time: int = 30) -> None:
        """
        发送消息并等待响应
        
        Args:
            message: 要发送的消息
            wait_time: 等待时间（秒）
        """
        print(f"\n发送消息: {message[:50]}{'...' if len(message) > 50 else ''}")
        self.send_message(message)
        print(f"等待 {wait_time} 秒...")
        time.sleep(wait_time)
        print("✓ 完成")


class KiroAutomation:
    """Kiro自动化主类 - 用于多轮对话"""
    
    def __init__(self, config: Optional[VisualConfig] = None):
        self.interface = KiroVisualInterface(config)
        self.message_history: List[str] = []
    
    def run_conversation(self, messages: List[str], wait_between: int = 30) -> None:
        """
        运行多轮对话
        
        Args:
            messages: 要发送的消息列表
            wait_between: 每条消息之间的等待时间
        """
        print(f"\n开始多轮对话，共 {len(messages)} 条消息")
        print("=" * 50)
        
        for i, msg in enumerate(messages, 1):
            print(f"\n[{i}/{len(messages)}] 发送消息...")
            self.interface.send_and_wait(msg, wait_between)
            self.message_history.append(msg)
        
        print("\n" + "=" * 50)
        print("✓ 对话完成")
    
    def run_iterative_task(
        self,
        initial_prompt: str,
        followup_template: str,
        max_iterations: int = 5,
        wait_time: int = 60
    ) -> None:
        """
        运行迭代任务
        
        Args:
            initial_prompt: 初始提示词
            followup_template: 后续提示词模板（可包含 {iteration} 占位符）
            max_iterations: 最大迭代次数
            wait_time: 每次迭代的等待时间
        """
        print(f"\n开始迭代任务，最多 {max_iterations} 次迭代")
        print("=" * 50)
        
        # 发送初始提示
        print("\n[初始] 发送初始提示...")
        self.interface.send_and_wait(initial_prompt, wait_time)
        
        # 迭代
        for i in range(1, max_iterations + 1):
            followup = followup_template.format(iteration=i)
            print(f"\n[迭代 {i}/{max_iterations}] 发送后续提示...")
            self.interface.send_and_wait(followup, wait_time)
        
        print("\n" + "=" * 50)
        print("✓ 迭代任务完成")


def calibrate_kiro():
    """校准工具入口"""
    interface = KiroVisualInterface()
    interface.calibrate()


def demo():
    """演示用法"""
    print("Kiro 视觉自动化演示")
    print("=" * 50)
    
    # 检查依赖
    if not HAS_VISUAL_DEPS:
        print("错误: 缺少依赖，请运行:")
        print("  pip install pyautogui pillow pyperclip")
        return
    
    # 加载或创建配置
    config = VisualConfig.load()
    
    if config.input_field_pos == (0, 0):
        print("首次运行，需要校准...")
        interface = KiroVisualInterface(config)
        interface.calibrate()
    else:
        print(f"已加载配置，输入框位置: {config.input_field_pos}")
    
    # 示例：发送单条消息
    automation = KiroAutomation(config)
    
    print("\n准备发送测试消息...")
    print("按 Ctrl+C 取消")
    
    try:
        time.sleep(3)
        automation.interface.send_message("你好，这是一条测试消息")
        print("✓ 消息已发送")
    except KeyboardInterrupt:
        print("\n已取消")


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == "calibrate":
        calibrate_kiro()
    else:
        demo()
