# -*- coding: utf-8 -*-
"""
Windows 图像查找与点击模块 - 增强版
支持窗口内查找、模板匹配、置信度设置
"""

import re
import time
import ctypes
import ctypes.wintypes
from pathlib import Path
from typing import Optional, Tuple, List, Callable

import pyautogui
import cv2
import numpy as np

# PyAutoGUI 安全设置
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.1

# Windows API 常量
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
WM_MOUSEMOVE = 0x0200

# 窗口状态常量
SW_RESTORE = 9
SW_SHOWMINIMIZED = 2
SW_SHOWMAXIMIZED = 3
SW_SHOWNOACTIVATE = 4
GWL_STYLE = -16
WS_VISIBLE = 0x10000000
WS_MINIMIZE = 0x20000000

# WINDOWPLACEMENT 结构体（ctypes.wintypes 没有这个，需要手动定义）
class _POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

class WINDOWPLACEMENT(ctypes.Structure):
    _fields_ = [
        ("length", ctypes.wintypes.UINT),
        ("flags", ctypes.wintypes.UINT),
        ("showCmd", ctypes.wintypes.UINT),
        ("ptMinPosition", _POINT),
        ("ptMaxPosition", _POINT),
        ("rcNormalPosition", ctypes.wintypes.RECT),
    ]


class WindowFinder:
    """Windows 窗口查找器"""

    def __init__(self):
        self._windows = []

    def find_window_by_title(self, title: str, exact: bool = False) -> Optional[int]:
        """
        通过标题查找窗口

        Args:
            title: 窗口标题（支持模糊匹配）
            exact: 是否精确匹配

        Returns:
            窗口句柄 (HWND) 或 None
        """
        matches = []

        def enum_callback(hwnd, _):
            if ctypes.windll.user32.IsWindowVisible(hwnd):
                length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
                if length > 0:
                    buffer = ctypes.create_unicode_buffer(length + 1)
                    ctypes.windll.user32.GetWindowTextW(hwnd, buffer, length + 1)
                    window_text = buffer.value

                    if exact:
                        if window_text == title:
                            matches.append((hwnd, window_text, len(window_text)))
                    else:
                        if title.lower() in window_text.lower():
                            matches.append((hwnd, window_text, len(window_text)))
            return True

        EnumWindowsProc = ctypes.WINFUNCTYPE(
            ctypes.wintypes.BOOL,
            ctypes.wintypes.HWND,
            ctypes.wintypes.LPARAM
        )
        ctypes.windll.user32.EnumWindows(EnumWindowsProc(enum_callback), 0)

        if not matches:
            return None

        # 优先选择精确匹配的窗口
        exact_matches = [m for m in matches if m[1] == title]
        if exact_matches:
            return exact_matches[0][0]

        # 否则选择标题最短的匹配（更可能是主窗口）
        matches.sort(key=lambda x: x[2])
        return matches[0][0]

    def find_all_windows(self) -> List[dict]:
        """获取所有可见窗口"""
        windows = []

        def enum_callback(hwnd, _):
            if ctypes.windll.user32.IsWindowVisible(hwnd):
                length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
                if length > 0:
                    buffer = ctypes.create_unicode_buffer(length + 1)
                    ctypes.windll.user32.GetWindowTextW(hwnd, buffer, length + 1)
                    windows.append({
                        'hwnd': hwnd,
                        'title': buffer.value
                    })
            return True

        EnumWindowsProc = ctypes.WINFUNCTYPE(
            ctypes.wintypes.BOOL,
            ctypes.wintypes.HWND,
            ctypes.wintypes.LPARAM
        )
        ctypes.windll.user32.EnumWindows(EnumWindowsProc(enum_callback), 0)
        return windows

    def get_window_rect(self, hwnd: int) -> Optional[Tuple[int, int, int, int]]:
        """获取窗口区域 (left, top, right, bottom)"""
        rect = ctypes.wintypes.RECT()
        if ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            return (rect.left, rect.top, rect.right, rect.bottom)
        return None

    def set_foreground(self, hwnd: int) -> bool:
        """激活窗口到前台"""
        return bool(ctypes.windll.user32.SetForegroundWindow(hwnd))

    def bring_to_front(self, hwnd: int) -> bool:
        """将窗口置顶"""
        SWP_NOMOVE = 0x0001
        SWP_NOSIZE = 0x0002
        HWND_TOP = 0
        return bool(ctypes.windll.user32.SetWindowPos(
            hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE
        ))

    def is_minimized(self, hwnd: int) -> bool:
        """判断窗口是否处于最小化状态"""
        placement = WINDOWPLACEMENT()
        placement.length = ctypes.sizeof(WINDOWPLACEMENT)
        if ctypes.windll.user32.GetWindowPlacement(hwnd, ctypes.byref(placement)):
            return placement.showCmd == SW_SHOWMINIMIZED
        return False

    def restore(self, hwnd: int) -> bool:
        """恢复（取消最小化）窗口"""
        # 先用 SW_RESTORE 恢复正常大小
        result1 = ctypes.windll.user32.ShowWindow(hwnd, SW_RESTORE)
        time.sleep(0.3)
        # 再激活并置前
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        time.sleep(0.2)
        # 额外用 SetWindowPos 确保置顶
        SWP_NOMOVE = 0x0001
        SWP_NOSIZE = 0x0002
        HWND_TOP = 0
        ctypes.windll.user32.SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE)
        return bool(result1)

    def is_window_too_small(self, hwnd: int, threshold: int = 50) -> bool:
        """判断窗口是否太小（可能是最小化状态）"""
        rect = self.get_window_rect(hwnd)
        if not rect:
            return True
        width = rect[2] - rect[0]
        height = rect[3] - rect[1]
        return width < threshold or height < threshold


class ImageClicker:
    """图像查找与点击器 - 增强版"""

    def __init__(
        self,
        confidence: float = 0.8,
        screenshot_delay: float = 0,
        window_title: Optional[str] = None
    ):
        """
        初始化

        Args:
            confidence: 匹配置信度 (0.0-1.0)
            screenshot_delay: 截图前延迟(秒)
            window_title: 限定窗口标题（可选）
        """
        self.confidence = confidence
        self.screenshot_delay = screenshot_delay
        self.window_finder = WindowFinder()
        self._target_window = None

        if window_title:
            self.set_window(window_title)

    def set_window(self, title: str, exact: bool = False) -> bool:
        """
        设置目标窗口

        Args:
            title: 窗口标题
            exact: 是否精确匹配

        Returns:
            是否成功找到窗口
        """
        hwnd = self.window_finder.find_window_by_title(title, exact)
        if hwnd:
            self._target_window = hwnd
            # 获取窗口标题
            length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
            buffer = ctypes.create_unicode_buffer(length + 1)
            ctypes.windll.user32.GetWindowTextW(hwnd, buffer, length + 1)
            self._window_title = buffer.value
            return True
        return False

    def _capture_window(self) -> Optional[np.ndarray]:
        """捕获目标窗口截图"""
        if not self._target_window:
            return None

        # 多次尝试激活窗口，确保它在前台
        for _ in range(3):
            self.window_finder.set_foreground(self._target_window)
            time.sleep(0.15)

        rect = self.window_finder.get_window_rect(self._target_window)
        if not rect:
            print(f"[ImageClicker] 无法获取窗口区域")
            return None

        left, top, right, bottom = rect
        width = right - left
        height = bottom - top

        # 跳过太小的窗口（可能是无效窗口）
        if width < 50 or height < 50:
            print(f"[ImageClicker] 窗口太小 ({width}x{height})")
            return None

        # 直接截取窗口区域（不含padding，避免边框干扰）
        try:
            screenshot = pyautogui.screenshot(
                region=(left, top, width, height)
            )
            img = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

            # 保存裁剪偏移量，用于计算原始坐标
            self._offset = (left, top)
            self._window_size = (width, height)

            return img

        except Exception as e:
            print(f"[ImageClicker] 截图失败: {e}")
            return None

    def _load_image(self, image_path: str) -> Optional[np.ndarray]:
        """加载图片文件（支持中文文件名，支持多目录查找）"""
        # 图片可能存放的位置列表
        search_paths = [
            Path.cwd() / image_path,                    # 1. 当前工作目录
            Path.cwd() / "images" / image_path,          # 2. images 子目录
            Path(__file__).parent / image_path,          # 3. 脚本同目录
            Path(__file__).parent / "images" / image_path, # 4. 脚本同目录/images
        ]

        found_path = None
        for sp in search_paths:
            if sp.exists():
                found_path = sp
                break

        if found_path is None:
            print(f"[ImageClicker] 文件不存在: {image_path}")
            print(f"  搜索路径: {[str(p) for p in search_paths]}")
            return None

        # OpenCV 不支持中文文件名，使用 Pillow 转换
        try:
            from PIL import Image
            pil_img = Image.open(found_path)
            img = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
            return img
        except Exception as e:
            print(f"[ImageClicker] 读取图片失败: {e}")
            return None

    def _take_screenshot(self) -> np.ndarray:
        """获取截图（窗口或全屏）"""
        if self.screenshot_delay > 0:
            time.sleep(self.screenshot_delay)

        if self._target_window:
            screen = self._capture_window()
            if screen is not None:
                return screen

        # 回退到全屏截图
        screenshot = pyautogui.screenshot()
        return cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

    def find_image(
        self,
        image_path: str,
        confidence: Optional[float] = None,
        screenshot: Optional[np.ndarray] = None
    ) -> Optional[Tuple[int, int]]:
        """
        在窗口/屏幕上查找图片位置

        Returns:
            原始屏幕坐标 (x, y) 或 None
        """
        template = self._load_image(image_path)
        if template is None:
            return None

        if screenshot is None:
            screen = self._take_screenshot()
        else:
            screen = screenshot

        conf = confidence if confidence is not None else self.confidence

        try:
            result = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

            if max_val >= conf:
                h, w = template.shape[:2]

                # 计算截图中的中心点
                center_x = max_loc[0] + w // 2
                center_y = max_loc[1] + h // 2

                # 如果是窗口截图，转换为屏幕坐标
                if hasattr(self, '_offset'):
                    offset_x, offset_y = self._offset
                    return (center_x + offset_x, center_y + offset_y)
                else:
                    return (center_x, center_y)

        except Exception as e:
            print(f"[ImageClicker] 匹配失败: {e}")

        return None

    def find_all_images(
        self,
        image_path: str,
        confidence: Optional[float] = None
    ) -> List[Tuple[int, int]]:
        """查找所有匹配的图片位置"""
        template = self._load_image(image_path)
        if template is None:
            return []

        screen = self._take_screenshot()
        conf = confidence if confidence is not None else self.confidence

        try:
            result = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
            locations = np.where(result >= conf)
            matches = []

            h, w = template.shape[:2]
            offset_x = getattr(self, '_offset', (0, 0))[0]
            offset_y = getattr(self, '_offset', (0, 0))[1]

            for pt in zip(*locations[::-1]):
                center_x = pt[0] + w // 2 + offset_x
                center_y = pt[1] + h // 2 + offset_y
                matches.append((center_x, center_y))

            return matches

        except Exception as e:
            print(f"[ImageClicker] 批量匹配失败: {e}")
            return []

    def click_image(
        self,
        image_path: str,
        confidence: Optional[float] = None,
        clicks: int = 1,
        button: str = 'left'
    ) -> bool:
        """
        查找并点击图片

        Returns:
            是否成功点击
        """
        pos = self.find_image(image_path, confidence)
        if pos is None:
            return False

        # 确保窗口在前台
        if self._target_window:
            self.window_finder.set_foreground(self._target_window)
            time.sleep(0.1)

        pyautogui.click(pos[0], pos[1], clicks=clicks, button=button)
        return True

    def wait_for_image(
        self,
        image_path: str,
        timeout: float = 10,
        interval: float = 0.5,
        confidence: Optional[float] = None
    ) -> Optional[Tuple[int, int]]:
        """等待图片出现"""
        start_time = time.time()

        while time.time() - start_time < timeout:
            pos = self.find_image(image_path, confidence)
            if pos:
                return pos
            time.sleep(interval)

        return None

    def click_when_visible(
        self,
        image_path: str,
        timeout: float = 10,
        interval: float = 0.5,
        confidence: Optional[float] = None,
        button: str = 'left'
    ) -> bool:
        """等待图片出现后点击"""
        pos = self.wait_for_image(image_path, timeout, interval, confidence)
        if pos:
            if self._target_window:
                self.window_finder.set_foreground(self._target_window)
                time.sleep(0.1)
            pyautogui.click(pos[0], pos[1], button=button)
            return True
        return False


class WindowImageClicker:
    """
    窗口图像点击器 - 简化用法
    自动查找窗口并在该窗口内查找图片
    """

    def __init__(self, window_title: str, confidence: float = 0.8, debug: bool = False):
        """
        初始化

        Args:
            window_title: 目标窗口标题
            confidence: 匹配置信度
            debug: 是否开启调试模式（保存截图）
        """
        self.window_title = window_title
        self.confidence = confidence
        self.debug = debug
        self.clicker = ImageClicker(confidence=confidence)
        self._window_handle = None

    def _ensure_window(self) -> bool:
        """确保窗口已激活"""
        if not self._window_handle:
            self._window_handle = self.clicker.window_finder.find_window_by_title(
                self.window_title
            )
            if not self._window_handle:
                print(f"[WindowImageClicker] 未找到窗口: {self.window_title}")
                return False
            self.clicker._target_window = self._window_handle

            # 获取并显示找到的窗口标题
            length = ctypes.windll.user32.GetWindowTextLengthW(self._window_handle)
            buffer = ctypes.create_unicode_buffer(length + 1)
            ctypes.windll.user32.GetWindowTextW(self._window_handle, buffer, length + 1)
            print(f"[WindowImageClicker] 已找到窗口: {buffer.value}")

        self.clicker.window_finder.set_foreground(self._window_handle)
        return True

    def find_and_click(self, image_path: str, timeout: float = 0) -> Tuple[bool, Optional[Tuple[int, int]]]:
        """
        查找窗口并在窗口内查找图片点击

        Args:
            image_path: 图片路径
            timeout: 等待图片出现的超时时间(0=立即)

        Returns:
            (是否成功, 图片位置)
        """
        if not self._ensure_window():
            print(f"[WindowImageClicker] 未找到窗口: {self.window_title}")
            return False, None

        # 获取窗口信息用于调试
        rect = self.clicker.window_finder.get_window_rect(self._window_handle)
        if rect:
            print(f"[调试] 窗口区域: 左={rect[0]}, 上={rect[1]}, 右={rect[2]}, 下={rect[3]}")
            print(f"[调试] 窗口尺寸: {rect[2]-rect[0]}x{rect[3]-rect[1]}")

        if timeout > 0:
            pos = self.clicker.wait_for_image(image_path, timeout=timeout)
        else:
            pos = self.clicker.find_image(image_path)

        if pos:
            self.clicker.window_finder.set_foreground(self._window_handle)
            time.sleep(0.1)
            pyautogui.click(pos[0], pos[1])
            return True, pos

        # 调试：保存截图
        if self.debug:
            self._save_debug_screenshot(image_path)

        return False, None

    def _save_debug_screenshot(self, image_path: str):
        """保存调试截图"""
        try:
            from pathlib import Path
            import os

            debug_dir = Path.cwd() / "debug_screenshots"
            debug_dir.mkdir(exist_ok=True)

            # 获取窗口截图
            self._ensure_window()
            window_img = self.clicker._capture_window()
            if window_img is not None:
                # 保存窗口截图
                debug_path = debug_dir / f"window_{int(time.time())}.png"
                cv2.imwrite(str(debug_path), window_img)
                print(f"[调试] 已保存窗口截图: {debug_path}")

            # 保存模板图片
            template = self.clicker._load_image(image_path)
            if template is not None:
                template_path = debug_dir / f"template_{image_path}"
                cv2.imwrite(str(template_path), template)
                print(f"[调试] 已保存模板图片: {template_path}")

        except Exception as e:
            print(f"[调试] 保存截图失败: {e}")

    def click(self, image_path: str, confidence: Optional[float] = None) -> bool:
        """
        快捷方法：查找并点击

        Args:
            image_path: 图片路径
            confidence: 匹配置信度

        Returns:
            是否成功
        """
        success, _ = self.find_and_click(image_path, timeout=0)
        return success


def create_workbuddy_clicker(confidence: float = 0.8, debug: bool = False) -> WindowImageClicker:
    """
    创建 WorkBuddy 窗口点击器

    Args:
        confidence: 匹配置信度
        debug: 是否开启调试模式

    Returns:
        WindowImageClicker 实例
    """
    return WindowImageClicker("WorkBuddy", confidence=confidence, debug=debug)


# ============== 便捷函数 ==============

def click_image(image_path: str, confidence: float = 0.8, window: Optional[str] = None) -> bool:
    """
    快速点击单张图片

    Args:
        image_path: 图片路径
        confidence: 匹配置信度
        window: 限定窗口标题（可选）

    Returns:
        是否成功
    """
    clicker = ImageClicker(confidence=confidence, window_title=window)
    return clicker.click_image(image_path)


def find_image_in_window(image_path: str, window_title: str, confidence: float = 0.8) -> Optional[Tuple[int, int]]:
    """在指定窗口中查找图片"""
    clicker = WindowImageClicker(window_title, confidence)
    success, pos = clicker.find_and_click(image_path)
    return pos if success else None


if __name__ == "__main__":
    print("=" * 50)
    print("Windows 窗口图像点击模块")
    print("=" * 50)
    print("\n使用示例:")
    print()
    print("  # 方式1: 在 WorkBuddy 窗口中点击")
    print("  clicker = create_workbuddy_clicker()")
    print("  clicker.click('自动化.png')")
    print()
    print("  # 方式2: 通用方式")
    print("  from image_clicker import WindowImageClicker")
    print("  clicker = WindowImageClicker('WorkBuddy')")
    print("  clicker.click('自动化.png')")
    print()
    print("  # 方式3: 等待图片出现后点击")
    print("  clicker.find_and_click('button.png', timeout=10)")


class SmartCodeExecutor:
    """
    智能代码执行器 - 自动处理 @image:xxx.png 标记
    """

    def __init__(self, image_clicker: Optional[ImageClicker] = None):
        self.clicker = image_clicker or ImageClicker()
        self.pattern = re.compile(r'@image:([^\s\)`]+)')

    def extract_image_refs(self, code: str) -> List[str]:
        """从代码中提取所有图像引用"""
        return self.pattern.findall(code)

    def process_image_refs(
        self,
        code: str,
        confidence: float = 0.8,
        click_delay: float = 0.5
    ) -> Tuple[List[Tuple[str, bool]], str]:
        """
        处理代码中的所有图像引用并执行点击

        Args:
            code: 包含 @image:xxx.png 的代码
            confidence: 匹配置信度
            click_delay: 每次点击后延迟

        Returns:
            (结果列表 [(图片路径, 是否成功), ...], 清理后的代码)
        """
        results = []
        clean_code = code

        for match in self.pattern.finditer(code):
            image_path = match.group(1)
            success = self.clicker.click_image(image_path, confidence=confidence)

            if success:
                print(f"[✓] 已点击: {image_path}")
            else:
                print(f"[✗] 未找到: {image_path}")

            results.append((image_path, success))
            time.sleep(click_delay)

        # 移除所有图像引用
        clean_code = self.pattern.sub('', clean_code)

        return results, clean_code

    def execute(self, code: str, confidence: float = 0.8) -> dict:
        """
        执行代码，自动处理图像引用

        Args:
            code: 包含 @image:xxx 的 Python 代码
            confidence: 匹配置信度

        Returns:
            包含执行结果的字典
        """
        results, clean_code = self.process_image_refs(code, confidence)

        return {
            'image_results': results,
            'clean_code': clean_code,
            'success_count': sum(1 for _, s in results if s),
            'fail_count': sum(1 for _, s in results if not s)
        }
