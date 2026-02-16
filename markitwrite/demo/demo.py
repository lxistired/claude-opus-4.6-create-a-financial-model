"""
markitwrite 演示脚本

用 repo 中已有的截图，分别粘贴到 Word、PPT、Markdown 三种格式，
生成实际文件供查看验证。
"""

import os
import sys

# 路径设置
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SCREENSHOT_DIR = os.path.join(REPO_ROOT, "screenshots")
DEMO_DIR = os.path.dirname(os.path.abspath(__file__))

# 选择输入图片
INPUT_IMAGES = [
    os.path.join(SCREENSHOT_DIR, "02-key-summary.png"),
    os.path.join(SCREENSHOT_DIR, "09-dcf.png"),
    os.path.join(SCREENSHOT_DIR, "10-sensitivity.png"),
]

from markitwrite import MarkItWrite

writer = MarkItWrite()


def demo_docx():
    """把多张截图粘贴进一个 Word 文档"""
    output_path = os.path.join(DEMO_DIR, "output_report.docx")

    # 第一张图 → 新建文档
    result = writer.paste(
        INPUT_IMAGES[0],
        target=output_path,
        size={"width": 6.0},
    )
    print(f"  [1/3] 粘贴 {os.path.basename(INPUT_IMAGES[0])} → 新建 Word")

    # 第二、三张图 → 追加到已有文档
    for i, img in enumerate(INPUT_IMAGES[1:], start=2):
        result = writer.paste(
            img,
            target=output_path,
            size={"width": 6.0},
        )
        print(f"  [{i}/3] 粘贴 {os.path.basename(img)} → 追加到 Word")

    fsize = os.path.getsize(output_path)
    print(f"  ✓ 生成: {output_path} ({fsize:,} bytes)")
    return output_path


def demo_pptx():
    """把每张截图分别粘贴到不同的幻灯片"""
    output_path = os.path.join(DEMO_DIR, "output_slides.pptx")

    # 第一张图 → 新建 PPT
    result = writer.paste(
        INPUT_IMAGES[0],
        target=output_path,
        size={"width": 8.0},
    )
    print(f"  [1/3] 粘贴 {os.path.basename(INPUT_IMAGES[0])} → 新建 PPT 幻灯片 1")

    # 后续图片 → 追加新幻灯片
    for i, img in enumerate(INPUT_IMAGES[1:], start=2):
        result = writer.paste(
            img,
            target=output_path,
            size={"width": 8.0},
        )
        print(f"  [{i}/3] 粘贴 {os.path.basename(img)} → PPT 幻灯片 {i}")

    fsize = os.path.getsize(output_path)
    print(f"  ✓ 生成: {output_path} ({fsize:,} bytes)")
    return output_path


def demo_markdown():
    """把截图以 base64 嵌入 Markdown 文件"""
    output_path = os.path.join(DEMO_DIR, "output_notes.md")

    # 先写入标题
    with open(output_path, "w") as f:
        f.write("# Financial Model Screenshots\n\n由 markitwrite 自动生成\n")

    # 逐张追加
    for i, img in enumerate(INPUT_IMAGES, start=1):
        result = writer.paste(
            img,
            target=output_path,
            embed=True,
            alt_text=os.path.basename(img),
        )
        print(f"  [{i}/3] 粘贴 {os.path.basename(img)} → Markdown (base64)")

    fsize = os.path.getsize(output_path)
    print(f"  ✓ 生成: {output_path} ({fsize:,} bytes)")
    return output_path


if __name__ == "__main__":
    # 检查输入图片是否存在
    for img in INPUT_IMAGES:
        if not os.path.isfile(img):
            print(f"错误: 找不到输入图片 {img}")
            sys.exit(1)

    print("=" * 60)
    print("markitwrite 演示 — 把截图粘贴到各种文档格式")
    print("=" * 60)
    print(f"\n输入图片: {len(INPUT_IMAGES)} 张来自 screenshots/ 目录\n")

    print("--- 1. 粘贴到 Word (.docx) ---")
    docx_path = demo_docx()

    print("\n--- 2. 粘贴到 PowerPoint (.pptx) ---")
    pptx_path = demo_pptx()

    print("\n--- 3. 粘贴到 Markdown (.md) ---")
    md_path = demo_markdown()

    print("\n" + "=" * 60)
    print("全部完成！生成的文件:")
    print(f"  Word:       {docx_path}")
    print(f"  PowerPoint: {pptx_path}")
    print(f"  Markdown:   {md_path}")
    print("=" * 60)
    print("\n请用 Word / PowerPoint / Markdown编辑器 打开这些文件验证图片是否正确显示。")
