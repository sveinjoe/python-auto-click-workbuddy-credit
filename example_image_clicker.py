# -*- coding: utf-8 -*-
"""
示例：自动点击 @image:xxx.png 格式的图像

使用前先安装依赖：
pip install -r requirements_image_clicker.txt
"""

import time
from image_clicker import ImageClicker, SmartCodeExecutor


def basic_example():
    """基础用法示例"""
    print("=" * 50)
    print("基础用法示例")
    print("=" * 50)

    clicker = ImageClicker()

    # 示例1：点击单张图片
    print("\n[示例1] 查找并点击 '自动化.png'")
    success = clicker.click_image("自动化.png")
    if success:
        print("✓ 点击成功")
    else:
        print("✗ 未找到图片")

    # 示例2：设置更高置信度
    print("\n[示例2] 高精度匹配 'ok_button.png' (confidence=0.95)")
    success = clicker.click_image("ok_button.png", confidence=0.95)
    print(f"结果: {'成功' if success else '未找到'}")

    # 示例3：等待图片出现
    print("\n[示例3] 等待 'loading_complete.png' (最多10秒)")
    pos = clicker.wait_for_image("loading_complete.png", timeout=10)
    print(f"结果: {pos if pos else '超时未找到'}")


def smart_code_example():
    """智能代码执行示例"""
    print("\n" + "=" * 50)
    print("智能代码执行示例")
    print("=" * 50)

    executor = SmartCodeExecutor()

    # 模拟包含图像引用的代码
    code = """
# 自动化测试脚本
@image:login_button.png  # 点击登录按钮
time.sleep(1)
@image:username_field.png  # 点击用户名输入框
type_username("admin")
@image:password_field.png  # 点击密码输入框
type_password("***")
@image:submit_button.png  # 点击提交
"""

    print("\n原始代码:")
    print(code)

    print("\n处理图像引用...")
    result = executor.execute(code)

    print(f"\n处理结果:")
    print(f"  成功: {result['success_count']} 个")
    print(f"  失败: {result['fail_count']} 个")

    print(f"\n清理后的代码:")
    print(result['clean_code'])


def batch_click_example():
    """批量点击示例"""
    print("\n" + "=" * 50)
    print("批量点击示例")
    print("=" * 50)

    clicker = ImageClicker()

    # 一组要点击的图片
    images = [
        "step1.png",
        "step2.png",
        "step3.png",
        "step4.png",
    ]

    print("\n按顺序点击多张图片...")
    for i, img in enumerate(images, 1):
        print(f"\n[步骤 {i}] 点击 {img}")
        success = clicker.click_image(img)

        if success:
            print(f"  ✓ 已点击")
        else:
            print(f"  ✗ 未找到，跳过")

        time.sleep(0.5)  # 等待UI响应


def main():
    """主函数"""
    print("\n" + "=" * 50)
    print("Windows 图像自动点击演示")
    print("=" * 50)
    print("\n请确保有一张名为 'test.png' 的图片在当前目录")
    print("按 Ctrl+C 退出\n")

    # 基础示例
    basic_example()

    # 智能代码示例
    smart_code_example()

    # 批量点击示例
    batch_click_example()

    print("\n" + "=" * 50)
    print("演示完成!")
    print("=" * 50)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n已退出")
