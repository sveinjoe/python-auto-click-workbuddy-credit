# Windows 图像查找与点击模块

基于 PyAutoGUI + OpenCV 实现，支持模板匹配和置信度设置。

## 安装依赖

```bash
pip install pyautogui opencv-python numpy Pillow
```

## 使用方法

```python
from image_clicker import ImageClicker

clicker = ImageClicker()

# 基础用法 - 查找并点击单张图片
result = clicker.click_image("image.png")

# 带置信度阈值（0.8 = 80% 相似度）
result = clicker.click_image("image.png", confidence=0.8)

# 查找但不确定点
pos = clicker.find_image("image.png")
if pos:
    print(f"图片位置: {pos}")

# 等待图片出现（最多等10秒，每0.5秒检查一次）
pos = clicker.wait_for_image("image.png", timeout=10)
```

## 与代码执行集成

在你的代码执行器中添加图像查找功能：

```python
import re
import time
from image_clicker import ImageClicker

class SmartCodeExecutor:
    def __init__(self):
        self.clicker = ImageClicker()

    def execute_with_image_detection(self, code: str):
        """
        执行代码时自动处理 @image:xxx.png 标记
        """
        # 找到所有图像引用
        pattern = r'@image:([^\s]+)'
        matches = re.findall(pattern, code)

        for image_path in matches:
            print(f"正在查找并点击: {image_path}")
            success = self.clicker.click_image(image_path, confidence=0.8)
            if success:
                print(f"✓ 已点击 {image_path}")
                time.sleep(0.5)  # 等待UI响应
            else:
                print(f"✗ 未找到 {image_path}")

        # 从代码中移除图像引用，执行剩余代码
        clean_code = re.sub(pattern, '', code)
        return clean_code
```

## 示例脚本

```python
# example_usage.py
from image_clicker import ImageClicker
import pyautogui
import time

def main():
    clicker = ImageClicker()

    print("5秒后开始查找并点击 'button.png'...")
    print("请将 'button.png' 放在屏幕可见位置")
    time.sleep(5)

    # 查找图片位置
    pos = clicker.find_image("button.png", confidence=0.8)

    if pos:
        print(f"找到图片位置: {pos}")
        clicker.click_image("button.png", confidence=0.8)
        print("已点击!")
    else:
        print("未找到图片")

if __name__ == "__main__":
    main()
```

## 配置说明

| 参数 | 说明 | 默认值 |
|------|------|--------|
| confidence | 匹配置信度 0-1 | 0.8 |
| screenshot_delay | 截图延迟(秒) | 0 |
| click_interval | 点击间隔(秒) | 0 |

## 注意事项

1. 图片需要与屏幕内容高度相似（灰色区域/模糊图片会降低匹配率）
2. 不同屏幕分辨率下图片可能需要重新截图
3. 建议使用 PNG 格式并裁剪到精确大小
4. 可配合 `pyautogui.locateOnScreen()` 调试查找效果
