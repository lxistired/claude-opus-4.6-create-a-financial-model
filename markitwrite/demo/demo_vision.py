#!/usr/bin/env python3
"""Vision Assistant Demo - 在本地机器上运行此脚本来测试完整流程。

前置条件:
    1. pip install 'markitwrite[vision]'   (安装 mss + openai)
    2. pip install 'markitwrite[all]'      (安装所有文档格式支持)
    3. export OPENROUTER_API_KEY="sk-or-..."
    4. 确保有一个显示器 (不能在无头服务器上截屏)

用法:
    # 方式1: 一句话完成全部
    python demo_vision.py

    # 方式2: CLI
    markitwrite assist "把屏幕上的表格截图放到report.docx里"

    # 方式3: 快速截屏 (不用AI)
    markitwrite quick -o screenshot.docx
"""

from __future__ import annotations

import os
import sys

# 确保能找到本地的 markitwrite
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))


def demo_1_full_pipeline():
    """演示1: 一句话 → 截屏 → AI识别 → 裁剪 → 粘贴到文档"""
    print("=" * 60)
    print("Demo 1: 一句话完整流程")
    print("=" * 60)

    from markitwrite.vision import VisionAssistant

    assistant = VisionAssistant(verbose=True)

    # 一句话搞定
    result = assistant.run(
        "截取屏幕上最重要的内容，放到demo_output.docx里"
    )

    print(f"\nResult: success={result.success}")
    print(f"Output: {result.output_path}")
    print(f"Summary: {result.summary}")
    print(f"Steps:")
    for step in result.steps:
        print(f"  [{step.status}] {step.step}: {step.detail} ({step.duration_ms}ms)")
    print()


def demo_2_with_existing_image():
    """演示2: 用现有图片 (不截屏) → AI识别 → 粘贴到PPT"""
    print("=" * 60)
    print("Demo 2: 用已有图片放到PPT")
    print("=" * 60)

    from markitwrite.vision import VisionAssistant

    # 用项目里已有的截图
    screenshots_dir = os.path.join(os.path.dirname(__file__), "..", "..", "screenshots")
    dcf_image = os.path.join(screenshots_dir, "09-dcf.png")

    if not os.path.exists(dcf_image):
        print(f"  跳过: 找不到 {dcf_image}")
        return

    assistant = VisionAssistant(verbose=True)

    result = assistant.run(
        "把这个DCF估值模型放到演示文稿的第1页",
        image_path=dcf_image,
        output_path="demo_dcf_slides.pptx",
    )

    print(f"\nResult: success={result.success}")
    print(f"Output: {result.output_path}")
    print(f"Summary: {result.summary}")
    print()


def demo_3_quick_capture():
    """演示3: 快速截全屏，不需要AI"""
    print("=" * 60)
    print("Demo 3: 快速截屏 (无AI)")
    print("=" * 60)

    from markitwrite.vision import VisionAssistant

    assistant = VisionAssistant(verbose=True)

    result = assistant.quick_capture(output_path="demo_quick.docx")

    print(f"\nResult: success={result.success}")
    print(f"Output: {result.output_path}")
    print(f"Summary: {result.summary}")
    print()


def demo_4_capture_region_only():
    """演示4: 只截图+定位，不粘贴 (返回裁剪后的图片字节)"""
    print("=" * 60)
    print("Demo 4: 只截图+AI定位 (不粘贴)")
    print("=" * 60)

    from markitwrite.vision import VisionAssistant

    assistant = VisionAssistant(verbose=True)

    cropped_bytes, region = assistant.capture_region(
        "找到屏幕上的浏览器地址栏"
    )

    if region:
        print(f"\nFound region: ({region.left}, {region.top}) {region.width}x{region.height}")
    else:
        print("\nNo specific region found, got full screenshot")

    print(f"Image size: {len(cropped_bytes)} bytes")

    # 保存裁剪后的图片
    with open("demo_cropped.png", "wb") as f:
        f.write(cropped_bytes)
    print(f"Saved to: demo_cropped.png")
    print()


def demo_5_components_standalone():
    """演示5: 单独使用各组件 (不走完整pipeline)"""
    print("=" * 60)
    print("Demo 5: 单独使用组件")
    print("=" * 60)

    # --- 5a: 单独截屏 ---
    print("\n5a: ScreenCapture 单独使用")
    from markitwrite.vision import ScreenCapture

    cap = ScreenCapture()
    img = cap.take_screenshot()
    print(f"  Screenshot: {img.width}x{img.height}")

    img_bytes = cap.image_to_bytes(img)
    print(f"  Bytes: {len(img_bytes)}")

    # --- 5b: 单独用 Vision 分析 ---
    print("\n5b: VisionAnalyzer 单独使用")
    from markitwrite.vision import VisionAnalyzer

    analyzer = VisionAnalyzer()
    result = analyzer.describe(img_bytes, "这个屏幕上有什么？用中文回答。")
    print(f"  Gemini says: {result.description[:200]}...")

    # --- 5c: 单独定位元素 ---
    print("\n5c: 定位屏幕元素")
    locate_result = analyzer.locate(
        image_bytes=img_bytes,
        target="任何文本编辑器或终端窗口",
        image_width=img.width,
        image_height=img.height,
    )
    if locate_result.regions:
        for r in locate_result.regions:
            print(f"  Found: {r['label']} at ({r['left']},{r['top']}) "
                  f"{r['width']}x{r['height']} conf={r['confidence']}")
    else:
        print(f"  Nothing found. Response: {locate_result.description[:100]}")
    print()


def demo_6_cli_examples():
    """打印 CLI 用法示例"""
    print("=" * 60)
    print("CLI 命令示例 (在终端里运行)")
    print("=" * 60)
    print()
    print("# 一句话截屏+粘贴:")
    print('  markitwrite assist "把屏幕上的DCF模型截图放到report.docx里"')
    print()
    print("# 用已有图片:")
    print('  markitwrite assist "把这张图放到PPT第2页" --image chart.png')
    print()
    print("# 指定输出:")
    print('  markitwrite assist "截取当前页面" -o capture.docx')
    print()
    print("# 快速截全屏 (不用AI):")
    print("  markitwrite quick -o screenshot.docx")
    print()
    print("# 截全屏到PPT:")
    print("  markitwrite quick -o screenshot.pptx --width 9.5")
    print()


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Vision Assistant Demo")
    parser.add_argument(
        "demo",
        nargs="?",
        default="all",
        choices=["1", "2", "3", "4", "5", "cli", "all"],
        help="Which demo to run (default: all)",
    )
    args = parser.parse_args()

    demos = {
        "1": demo_1_full_pipeline,
        "2": demo_2_with_existing_image,
        "3": demo_3_quick_capture,
        "4": demo_4_capture_region_only,
        "5": demo_5_components_standalone,
        "cli": demo_6_cli_examples,
    }

    if args.demo == "all":
        # 先打印 CLI 用法
        demo_6_cli_examples()

        # 检查环境
        print("=" * 60)
        print("环境检查")
        print("=" * 60)

        checks = [
            ("mss (截屏)", "mss"),
            ("openai (OpenRouter API)", "openai"),
            ("python-docx (Word)", "docx"),
            ("python-pptx (PPT)", "pptx"),
            ("Pillow (图像处理)", "PIL"),
        ]
        all_ok = True
        for name, module in checks:
            try:
                __import__(module)
                print(f"  [OK] {name}")
            except ImportError:
                print(f"  [MISSING] {name}")
                all_ok = False

        api_key = os.environ.get("OPENROUTER_API_KEY", "")
        if api_key:
            print(f"  [OK] OPENROUTER_API_KEY (set, ...{api_key[-4:]})")
        else:
            print("  [MISSING] OPENROUTER_API_KEY - AI功能不可用")
            all_ok = False

        if not all_ok:
            print("\n安装缺失依赖:")
            print("  pip install 'markitwrite[all]'")
            print("  export OPENROUTER_API_KEY='sk-or-...'")
            print()

        # 运行 demo 2 (不需要截屏，最安全)
        print("\n运行 Demo 2 (用已有图片)...\n")
        demo_2_with_existing_image()
    else:
        demos[args.demo]()
