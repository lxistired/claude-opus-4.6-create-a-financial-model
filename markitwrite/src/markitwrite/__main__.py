"""CLI entry point for markitwrite.

Usage:
    markitwrite paste screenshot.png --to output.docx
    markitwrite paste chart.png --to slides.pptx --slide 3
    markitwrite paste diagram.png --to notes.md --embed
    markitwrite paste image.png --to report.docx --width 4 --height 3
    markitwrite formats
"""

from __future__ import annotations

import argparse
import sys


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="markitwrite",
        description="AI virtual clipboard - paste images into any document format.",
    )
    subparsers = parser.add_subparsers(dest="command")

    # -- paste command --
    paste_parser = subparsers.add_parser(
        "paste", help="Paste an image into a document."
    )
    paste_parser.add_argument("image", help="Path to the image file.")
    paste_parser.add_argument(
        "--to", dest="target", required=True, help="Target document path."
    )
    paste_parser.add_argument(
        "--format",
        dest="target_format",
        default=None,
        help="Explicit target format (e.g. .docx). Inferred from --to if omitted.",
    )
    paste_parser.add_argument(
        "--width", type=float, default=None, help="Image width in inches."
    )
    paste_parser.add_argument(
        "--height", type=float, default=None, help="Image height in inches."
    )
    paste_parser.add_argument(
        "--slide", type=int, default=None, help="Slide number for PPTX (1-indexed)."
    )
    paste_parser.add_argument(
        "--paragraph",
        type=int,
        default=None,
        help="Paragraph index for DOCX (0-indexed).",
    )
    paste_parser.add_argument(
        "--embed",
        action="store_true",
        default=True,
        help="Embed image as base64 in Markdown (default).",
    )
    paste_parser.add_argument(
        "--no-embed",
        action="store_true",
        help="Use file reference instead of base64 for Markdown.",
    )
    paste_parser.add_argument(
        "--alt", default="image", help="Alt text for the image."
    )

    # -- formats command --
    subparsers.add_parser("formats", help="List supported document formats.")

    # -- assist command (vision assistant) --
    assist_parser = subparsers.add_parser(
        "assist",
        help="AI vision assistant: one sentence → screenshot → paste into document.",
    )
    assist_parser.add_argument(
        "instruction",
        help='Natural language instruction, e.g. "把屏幕上的DCF模型截图放到report.docx里"',
    )
    assist_parser.add_argument(
        "--image",
        default=None,
        help="Use this image instead of taking a screenshot.",
    )
    assist_parser.add_argument(
        "--output", "-o",
        default=None,
        help="Override output document path (inferred from instruction if omitted).",
    )
    assist_parser.add_argument(
        "--monitor",
        type=int,
        default=0,
        help="Monitor index for screenshot (0=all, 1=first, etc.)",
    )
    assist_parser.add_argument(
        "--model",
        default=None,
        help="Claude model to use for vision analysis.",
    )
    assist_parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="Suppress progress output.",
    )

    # -- quick command (screenshot + paste, no AI) --
    quick_parser = subparsers.add_parser(
        "quick",
        help="Quick capture: screenshot full screen and paste into document (no AI).",
    )
    quick_parser.add_argument(
        "--output", "-o",
        default="output.docx",
        help="Output document path (default: output.docx).",
    )
    quick_parser.add_argument(
        "--monitor",
        type=int,
        default=0,
        help="Monitor index (0=all, 1=first, etc.)",
    )
    quick_parser.add_argument(
        "--width",
        type=float,
        default=None,
        help="Image width in inches.",
    )

    args = parser.parse_args(argv)

    if args.command is None:
        parser.print_help()
        return 1

    # Lazy import to avoid loading deps for --help
    from markitwrite import MarkItWrite

    writer = MarkItWrite()

    if args.command == "formats":
        formats = writer.supported_formats()
        if formats:
            print("Supported formats:")
            for fmt in formats:
                print(f"  {fmt}")
        else:
            print("No formats available. Install optional dependencies:")
            print("  pip install 'markitwrite[all]'")
        return 0

    if args.command == "paste":
        # Build position dict
        position = {}
        if args.slide is not None:
            position["slide"] = args.slide
        if args.paragraph is not None:
            position["paragraph"] = args.paragraph

        # Build size dict
        size = {}
        if args.width is not None:
            size["width"] = args.width
        if args.height is not None:
            size["height"] = args.height

        embed = not args.no_embed

        try:
            result = writer.paste(
                image_source=args.image,
                target=args.target,
                target_format=args.target_format,
                position=position or None,
                size=size or None,
                embed=embed,
                alt_text=args.alt,
            )
            print(
                f"Done: pasted image into {args.target} "
                f"({result.target_format}, {len(result.output)} bytes)"
            )
            return 0
        except Exception as e:
            print(f"Error: {e}", file=sys.stderr)
            return 1

    if args.command == "assist":
        try:
            from markitwrite.vision import VisionAssistant

            assistant = VisionAssistant(
                model=args.model,
                verbose=not args.quiet,
            )
            result = assistant.run(
                instruction=args.instruction,
                image_path=args.image,
                output_path=args.output,
                monitor=args.monitor,
            )
            if result.success:
                print(f"Done: {result.summary}")
                return 0
            else:
                print(f"Failed: {result.summary}", file=sys.stderr)
                return 1
        except ImportError as e:
            print(
                f"Error: Vision assistant requires extra dependencies.\n"
                f"  pip install 'markitwrite[vision]'\n"
                f"Detail: {e}",
                file=sys.stderr,
            )
            return 1
        except Exception as e:
            print(f"Error: {e}", file=sys.stderr)
            return 1

    if args.command == "quick":
        try:
            from markitwrite.vision import VisionAssistant

            assistant = VisionAssistant(verbose=True)
            size = {"width": args.width} if args.width else None
            result = assistant.quick_capture(
                output_path=args.output,
                monitor=args.monitor,
                size=size,
            )
            if result.success:
                print(f"Done: {result.summary}")
                return 0
            else:
                print(f"Failed: {result.summary}", file=sys.stderr)
                return 1
        except ImportError as e:
            print(
                f"Error: Quick capture requires 'mss' or 'pyautogui'.\n"
                f"  pip install mss\n"
                f"Detail: {e}",
                file=sys.stderr,
            )
            return 1
        except Exception as e:
            print(f"Error: {e}", file=sys.stderr)
            return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
