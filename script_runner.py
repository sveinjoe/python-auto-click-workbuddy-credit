# -*- coding: utf-8 -*-
"""
简易自动化脚本解析器
支持自然语言风格的脚本命令
"""

import re
import time
import sys
import ctypes
import ctypes.wintypes
from pathlib import Path

from image_clicker import (
    ImageClicker,
    WindowImageClicker,
    WindowFinder,
    create_workbuddy_clicker
)


class ScriptRunner:
    """自动化脚本运行器"""

    # 支持的命令模式
    COMMANDS = {
        # 找窗口
        r'^找窗口[：:]\s*(.+)$': 'find_window',
        r'^切换窗口[：:]\s*(.+)$': 'find_window',
        r'^窗口[：:]\s*(.+)$': 'find_window',

        # 点击图片
        r'^点击图片[：:]\s*(.+?)(?:\s*,\s*置信度\s*[=:]\s*[\d.]+)?$': 'click_image',
        r'^点击[：:]\s*(.+?)(?:\s*,\s*置信度\s*[=:]\s*[\d.]+)?$': 'click_image',
        r'^点击\s+(.+)$': 'click_image',

        # 右键点击图片
        r'^右键点击图片[：:]\s*(.+)$': 'right_click_image',
        r'^右键[：:]\s*(.+)$': 'right_click_image',

        # 双击图片
        r'^双击图片[：:]\s*(.+)$': 'double_click_image',
        r'^双击[：:]\s*(.+)$': 'double_click_image',

        # 鼠标移动
        r'^鼠标移动到[：:]\s*(\d+)\s*[，,]\s*(\d+)$': 'move_to',
        r'^移动到[：:]\s*(\d+)\s*[，,]\s*(\d+)$': 'move_to',
        r'^坐标[：:]\s*(\d+)\s*[，,]\s*(\d+)$': 'move_to',

        # 鼠标相对移动
        r'^鼠标移动[：:]\s*([+-]?\d+)\s*[，,]\s*([+-]?\d+)$': 'move_relative',
        r'^相对移动[：:]\s*([+-]?\d+)\s*[，,]\s*([+-]?\d+)$': 'move_relative',

        # 点击当前位置
        r'^点击$': 'click_current',
        r'^左键$': 'click_current',
        r'^单击$': 'click_current',

        # 右键当前位置
        r'^右键$': 'right_click_current',
        r'^右击$': 'right_click_current',

        # 双击当前位置
        r'^双击$': 'double_click_current',

        # 滚动
        r'^滚轮[上向下]*[：:]\s*(\d+)$': 'scroll',
        r'^滚轮[：:]\s*([+-]?\d+)$': 'scroll',
        r'^滚动[：:]\s*(\d+)$': 'scroll',

        # 等待
        r'^等待[：:]\s*(\d+(?:\.\d+)?)\s*(?:秒)?$': 'wait',
        r'^\[等待(\d+(?:\.\d+)?)\]$': 'wait',

        # 输入文本
        r'^输入[：:]\s*(.+)$': 'type_text',
        r'^打字[：:]\s*(.+)$': 'type_text',

        # 按键
        r'^按键[：:]\s*(.+)$': 'press_key',
        r'^按[：:]\s*(.+)$': 'press_key',

        # 按住/释放
        r'^按住[：:]\s*(.+)$': 'key_down',
        r'^释放[：:]\s*(.+)$': 'key_up',

        # 截图
        r'^截图[到:]\s*(.+)$': 'screenshot',

        # 注释
        r'^#': 'comment',
        r'^//': 'comment',
        r'^$': 'skip',  # 空行

        # 调试模式
        r'^调试[：:]\s*(开|关|on|off)$': 'debug_mode',

        # 置信度设置
        r'^置信度[：:]\s*([\d.]+)$': 'set_confidence',
    }

    def __init__(self, default_confidence: float = 0.8):
        """
        初始化

        Args:
            default_confidence: 默认匹配置信度
        """
        self.default_confidence = default_confidence
        self.current_confidence = default_confidence
        self.current_window = None
        self.current_clicker = None
        self.debug = False
        self._last_image_found = None

        # 编译所有正则
        self._patterns = []
        for pattern, cmd in self.COMMANDS.items():
            self._patterns.append((re.compile(pattern, re.IGNORECASE), cmd))

    def _log(self, msg: str):
        """打印日志"""
        if self.debug:
            print(f"  [DEBUG] {msg}")
        else:
            print(msg)

    def _find_command(self, line: str):
        """查找匹配的命令"""
        for pattern, cmd in self._patterns:
            match = pattern.match(line.strip())
            if match:
                return cmd, match
        return None, None

    def _get_clicker(self) -> ImageClicker:
        """获取或创建点击器"""
        if self.current_clicker is None:
            self.current_clicker = ImageClicker(confidence=self.current_confidence)
        self.current_clicker.confidence = self.current_confidence
        return self.current_clicker

    def run_line(self, line: str) -> bool:
        """
        执行单行脚本

        Args:
            line: 脚本行

        Returns:
            是否成功
        """
        line = line.strip()
        if not line:
            return True

        cmd, match = self._find_command(line)
        if cmd is None:
            self._log(f"❓ 未知命令: {line}")
            return False

        # 执行命令
        try:
            if cmd == 'find_window':
                return self.cmd_find_window(match.group(1))

            elif cmd == 'click_image':
                return self.cmd_click_image(match.group(1))

            elif cmd == 'right_click_image':
                return self.cmd_right_click_image(match.group(1))

            elif cmd == 'double_click_image':
                return self.cmd_double_click_image(match.group(1))

            elif cmd == 'move_to':
                return self.cmd_move_to(int(match.group(1)), int(match.group(2)))

            elif cmd == 'move_relative':
                return self.cmd_move_relative(int(match.group(1)), int(match.group(2)))

            elif cmd == 'click_current':
                return self.cmd_click_current()

            elif cmd == 'right_click_current':
                return self.cmd_right_click_current()

            elif cmd == 'double_click_current':
                return self.cmd_double_click_current()

            elif cmd == 'scroll':
                return self.cmd_scroll(int(match.group(1)))

            elif cmd == 'wait':
                return self.cmd_wait(float(match.group(1)))

            elif cmd == 'type_text':
                return self.cmd_type_text(match.group(1))

            elif cmd == 'press_key':
                return self.cmd_press_key(match.group(1))

            elif cmd == 'key_down':
                return self.cmd_key_down(match.group(1))

            elif cmd == 'key_up':
                return self.cmd_key_up(match.group(1))

            elif cmd == 'screenshot':
                return self.cmd_screenshot(match.group(1))

            elif cmd == 'debug_mode':
                return self.cmd_debug_mode(match.group(1))

            elif cmd == 'set_confidence':
                return self.cmd_set_confidence(float(match.group(1)))

            elif cmd in ('comment', 'skip'):
                return True

        except Exception as e:
            self._log(f"❌ 执行失败: {e}")
            return False

        return True

    # ==================== 命令实现 ====================

    def cmd_find_window(self, title: str) -> bool:
        """查找窗口"""
        self._log(f"🔍 查找窗口: {title}")

        finder = WindowFinder()
        hwnd = finder.find_window_by_title(title)

        if hwnd:
            # 获取窗口标题
            length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
            buffer = ctypes.create_unicode_buffer(length + 1)
            ctypes.windll.user32.GetWindowTextW(hwnd, buffer, length + 1)
            window_title = buffer.value

            # 获取窗口区域
            rect = finder.get_window_rect(hwnd)

            self.current_window = {
                'hwnd': hwnd,
                'title': window_title,
                'rect': rect
            }
            self.current_clicker = None  # 重置点击器

            self._log(f"✅ 已切换到窗口: {window_title}")
            if rect:
                self._log(f"   位置: ({rect[0]}, {rect[1]}) - ({rect[2]}, {rect[3]})")
            return True
        else:
            self._log(f"❌ 未找到窗口: {title}")
            return False

    def cmd_click_image(self, image_name: str) -> bool:
        """点击图片"""
        self._log(f"🖱️ 点击图片: {image_name} (置信度: {self.current_confidence})")

        clicker = self._get_clicker()

        # 添加 .png 后缀（如果需要）
        if not image_name.endswith('.png'):
            image_name = image_name + '.png'

        # 如果有当前窗口，设置为目标窗口
        if self.current_window:
            clicker._target_window = self.current_window['hwnd']
            # 激活目标窗口
            ctypes.windll.user32.SetForegroundWindow(self.current_window['hwnd'])
            time.sleep(0.15)

        pos = clicker.find_image(image_name)

        if pos:
            self._last_image_found = pos
            import pyautogui
            pyautogui.click(pos[0], pos[1])
            self._log(f"✅ 已点击位置: {pos}")
            return True
        else:
            self._log(f"❌ 未找到图片: {image_name}")
            return False

    def cmd_right_click_image(self, image_name: str) -> bool:
        """右键点击图片"""
        self._log(f"🖱️ 右键点击图片: {image_name}")

        clicker = self._get_clicker()

        if not image_name.endswith('.png'):
            image_name = image_name + '.png'

        if self.current_window:
            clicker._target_window = self.current_window['hwnd']

        pos = clicker.find_image(image_name)

        if pos:
            self._last_image_found = pos
            clicker.window_finder.set_foreground(self.current_window['hwnd'] if self.current_window else None)
            time.sleep(0.1)
            import pyautogui
            pyautogui.click(pos[0], pos[1], button='right')
            self._log(f"✅ 已右键点击位置: {pos}")
            return True
        else:
            self._log(f"❌ 未找到图片: {image_name}")
            return False

    def cmd_double_click_image(self, image_name: str) -> bool:
        """双击图片"""
        self._log(f"🖱️ 双击图片: {image_name}")

        clicker = self._get_clicker()

        if not image_name.endswith('.png'):
            image_name = image_name + '.png'

        if self.current_window:
            clicker._target_window = self.current_window['hwnd']

        pos = clicker.find_image(image_name)

        if pos:
            self._last_image_found = pos
            clicker.window_finder.set_foreground(self.current_window['hwnd'] if self.current_window else None)
            time.sleep(0.1)
            import pyautogui
            pyautogui.doubleClick(pos[0], pos[1])
            self._log(f"✅ 已双击位置: {pos}")
            return True
        else:
            self._log(f"❌ 未找到图片: {image_name}")
            return False

    def cmd_move_to(self, x: int, y: int) -> bool:
        """移动鼠标到坐标"""
        self._log(f"🖱️ 移动鼠标到: ({x}, {y})")

        import pyautogui
        pyautogui.moveTo(x, y, duration=0.2)
        self._log(f"✅ 鼠标已移动")
        return True

    def cmd_move_relative(self, dx: int, dy: int) -> bool:
        """相对移动鼠标"""
        self._log(f"🖱️ 相对移动鼠标: ({dx:+d}, {dy:+d})")

        import pyautogui
        current = pyautogui.position()
        new_x = current.x + dx
        new_y = current.y + dy
        pyautogui.moveTo(new_x, new_y, duration=0.2)
        self._log(f"✅ 鼠标已移动到: ({new_x}, {new_y})")
        return True

    def cmd_click_current(self) -> bool:
        """点击当前位置"""
        self._log(f"🖱️ 点击当前位置")
        import pyautogui
        pyautogui.click()
        self._log(f"✅ 已点击")
        return True

    def cmd_right_click_current(self) -> bool:
        """右键点击当前位置"""
        self._log(f"🖱️ 右键点击当前位置")
        import pyautogui
        pyautogui.click(button='right')
        self._log(f"✅ 已右键点击")
        return True

    def cmd_double_click_current(self) -> bool:
        """双击当前位置"""
        self._log(f"🖱️ 双击当前位置")
        import pyautogui
        pyautogui.doubleClick()
        self._log(f"✅ 已双击")
        return True

    def cmd_scroll(self, amount: int) -> bool:
        """滚动鼠标"""
        self._log(f"🖱️ 滚轮滚动: {amount}")

        import pyautogui
        pyautogui.scroll(amount)
        self._log(f"✅ 已滚动")
        return True

    def cmd_wait(self, seconds: float) -> bool:
        """等待"""
        self._log(f"⏳ 等待 {seconds} 秒...")
        time.sleep(seconds)
        self._log(f"✅ 等待完成")
        return True

    def cmd_type_text(self, text: str) -> bool:
        """输入文本（支持中文，使用 Windows 剪贴板粘贴）"""
        self._log(f"⌨️ 输入文本: {text}")

        import pyautogui

        # 确保焦点在目标窗口
        if self.current_window:
            hwnd = self.current_window['hwnd']
            self._log(f"   激活目标窗口: {self.current_window['title']}")
            ctypes.windll.user32.SetForegroundWindow(hwnd)
            time.sleep(0.3)

            # 点击窗口内部，确保焦点在输入框
            rect = self.current_window.get('rect')
            if rect:
                center_x = (rect[0] + rect[2]) // 2
                center_y = (rect[1] + rect[3]) // 2
                pyautogui.click(center_x, center_y)
                time.sleep(0.2)
        else:
            self._log(f"   ⚠️ 未设置目标窗口，可能输入到错误位置")

        # 设置剪贴板并粘贴
        if self._set_clipboard_text(text):
            time.sleep(0.1)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.1)
            self._log(f"✅ 输入完成")
            return True
        else:
            self._log(f"❌ 剪贴板设置失败")
            return False

    def _set_clipboard_text(self, text: str) -> bool:
        """设置剪贴板文本（通过临时文件，最可靠）"""
        import subprocess
        import tempfile
        import os

        temp_path = None
        try:
            # 写入 UTF-8 临时文件
            fd, temp_path = tempfile.mkstemp(suffix='.txt')
            with os.fdopen(fd, 'w', encoding='utf-8') as f:
                f.write(text)

            # PowerShell 读取文件并设置剪贴板
            cmd = ['powershell', '-NoProfile', '-Command',
                   f"Get-Content -Path '{temp_path}' -Encoding UTF8 | Set-Clipboard"]

            result = subprocess.run(cmd, capture_output=True, text=True, timeout=5)

            if result.returncode == 0:
                return True
            else:
                self._log(f"剪贴板命令失败: {result.stderr.strip()}")
                return False

        except Exception as e:
            self._log(f"剪贴板设置失败: {e}")
            return False
        finally:
            # 清理临时文件
            if temp_path and os.path.exists(temp_path):
                try:
                    os.unlink(temp_path)
                except:
                    pass

    def cmd_press_key(self, key: str) -> bool:
        """按键"""
        self._log(f"⌨️ 按键: {key}")

        import pyautogui
        pyautogui.press(key)
        self._log(f"✅ 按键完成")
        return True

    def cmd_key_down(self, key: str) -> bool:
        """按住按键"""
        self._log(f"⌨️ 按住: {key}")

        import pyautogui
        pyautogui.keyDown(key)
        self._log(f"✅ 已按住")
        return True

    def cmd_key_up(self, key: str) -> bool:
        """释放按键"""
        self._log(f"⌨️ 释放: {key}")

        import pyautogui
        pyautogui.keyUp(key)
        self._log(f"✅ 已释放")
        return True

    def cmd_screenshot(self, filename: str) -> bool:
        """截图"""
        self._log(f"📸 截图保存到: {filename}")

        import pyautogui
        img = pyautogui.screenshot()
        img.save(filename)
        self._log(f"✅ 截图已保存")
        return True

    def cmd_debug_mode(self, mode: str) -> bool:
        """切换调试模式"""
        self.debug = mode.lower() in ('开', 'on', 'true')
        self._log(f"🔧 调试模式: {'开启' if self.debug else '关闭'}")
        return True

    def cmd_set_confidence(self, confidence: float) -> bool:
        """设置置信度"""
        self.current_confidence = confidence
        self._log(f"📊 置信度已设置为: {confidence}")
        return True

    def run_script(self, script: str) -> dict:
        """
        执行脚本

        Args:
            script: 脚本内容（多行）

        Returns:
            执行结果统计
        """
        lines = script.strip().split('\n')

        success_count = 0
        fail_count = 0
        results = []

        print("=" * 50)
        print("开始执行脚本")
        print("=" * 50)

        for i, line in enumerate(lines, 1):
            # 跳过注释和空行
            stripped = line.strip()
            if not stripped or stripped.startswith('#') or stripped.startswith('//'):
                continue

            print(f"\n[行 {i}] {line}")

            success = self.run_line(line)

            if success:
                success_count += 1
            else:
                fail_count += 1

            results.append({
                'line': i,
                'content': line,
                'success': success
            })

        print("\n" + "=" * 50)
        print("脚本执行完成")
        print("=" * 50)
        print(f"✅ 成功: {success_count}")
        print(f"❌ 失败: {fail_count}")

        return {
            'success_count': success_count,
            'fail_count': fail_count,
            'results': results
        }

    def run_file(self, filepath: str) -> dict:
        """从文件执行脚本"""
        path = Path(filepath)
        if not path.exists():
            print(f"❌ 文件不存在: {filepath}")
            return {'success_count': 0, 'fail_count': 1, 'results': []}

        with open(path, 'r', encoding='utf-8') as f:
            script = f.read()

        return self.run_script(script)


# ==================== 便捷函数 ====================

def run(script: str) -> dict:
    """执行脚本字符串"""
    runner = ScriptRunner()
    return runner.run_script(script)


def run_file(filepath: str) -> dict:
    """执行脚本文件"""
    runner = ScriptRunner()
    return runner.run_file(filepath)


# ==================== 主程序 ====================

if __name__ == "__main__":
    import sys

    print("=" * 50)
    print("简易自动化脚本解析器")
    print("=" * 50)

    if len(sys.argv) > 1:
        # 从文件执行
        filepath = sys.argv[1]
        run_file(filepath)
    else:
        # 交互模式
        print("\n请输入脚本命令（输入 'exit' 或 'quit' 退出）:")
        print("-" * 40)

        runner = ScriptRunner(debug=True)

        while True:
            try:
                line = input("\n> ")
            except (KeyboardInterrupt, EOFError):
                print("\n\n再见!")
                break

            if line.lower() in ('exit', 'quit', 'q'):
                print("再见!")
                break

            if not line.strip():
                continue

            runner.run_line(line)
