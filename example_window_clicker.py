# -*- coding: utf-8 -*-
"""
示例：查找 WorkBuddy 窗口并点击图片

使用前先安装依赖：
pip install pyautogui opencv-python numpy Pillow
"""

import time
from image_clicker import (
    WindowImageClicker,
    create_workbuddy_clicker,
    ImageClicker,
    WindowFinder
)


def basic_example():
    """基础用法 - 在 WorkBuddy 窗口中点击"""
    print("=" * 50)
    print("基础用法")
    print("=" * 50)

    # 创建 WorkBuddy 点击器（开启调试模式）
    clicker = create_workbuddy_clicker(confidence=0.8, debug=True)

    # 查找并点击图片
    print("\n正在 WorkBuddy 窗口中查找 '自动化.png'...")

    success = clicker.click("自动化.png")

    if success:
        print("✓ 点击成功!")
    else:
        print("✗ 未找到图片")
        print("\n提示：调试截图已保存到 debug_screenshots/ 目录")


def wait_and_click_example():
    """等待图片出现后点击"""
    print("\n" + "=" * 50)
    print("等待图片出现")
    print("=" * 50)

    clicker = create_workbuddy_clicker()

    print("\n等待 '自动化.png' 出现（最多等待30秒）...")
    print("提示: 请确保图片在 WorkBuddy 窗口中可见")

    success, pos = clicker.find_and_click("自动化.png", timeout=30)

    if success:
        print(f"✓ 找到并点击! 位置: {pos}")
    else:
        print("✗ 超时未找到图片")


def list_windows_example():
    """列出所有窗口"""
    print("\n" + "=" * 50)
    print("查看所有窗口")
    print("=" * 50)

    finder = WindowFinder()
    windows = finder.find_all_windows()

    print(f"\n找到 {len(windows)} 个窗口:\n")
    for i, win in enumerate(windows, 1):
        print(f"  {i}. {win['title']}")


def custom_window_example():
    """自定义窗口"""
    print("\n" + "=" * 50)
    print("自定义窗口")
    print("=" * 50)

    # 查找任意窗口
    clicker = ImageClicker()

    # 模糊匹配
    hwnd = clicker.window_finder.find_window_by_title("Notepad")
    if hwnd:
        print(f"\n找到 Notepad 窗口: {hwnd}")
        clicker.set_window_by_handle(hwnd)
    else:
        print("\n未找到 Notepad 窗口")


def advanced_example():
    """高级用法 - 多次点击"""
    print("\n" + "=" * 50)
    print("高级用法 - 批量操作")
    print("=" * 50)

    clicker = create_workbuddy_clicker()

    images = [
        "自动化.png",
        "button1.png",
        "button2.png",
    ]

    print("\n按顺序点击多张图片...")
    for i, img in enumerate(images, 1):
        print(f"\n[步骤 {i}] 点击 {img}")
        success = clicker.click(img)

        if success:
            print(f"  ✓ 成功")
        else:
            print(f"  ✗ 未找到")

        time.sleep(0.5)


def main():
    print("\n" + "=" * 60)
    print("  Windows 窗口图像自动点击演示")
    print("=" * 60)
    print("\n本脚本演示如何:")
    print("  1. 查找标题为 'WorkBuddy' 的窗口")
    print("  2. 在该窗口内查找 '自动化.png'")
    print("  3. 点击找到的图片")
    print("\n" + "-" * 60)

    # 列出所有窗口
    list_windows_example()

    # 基础示例
    basic_example()

    print("\n" + "=" * 60)
    print("演示完成!")
    print("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n已退出")
