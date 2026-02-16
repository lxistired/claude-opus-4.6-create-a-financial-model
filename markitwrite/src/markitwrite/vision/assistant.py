"""VisionAssistant - one sentence in, done.

The orchestrator that ties everything together:
  capture (eyes) + analyzer (brain) + markitwrite (hands)

Usage:
    assistant = VisionAssistant()
    result = assistant.run("把屏幕上的DCF模型截图放到report.docx里")
    # Done. Screenshot taken, cropped, pasted into report.docx.
"""

from __future__ import annotations

import os
import sys
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Optional

from PIL import Image

from markitwrite.vision.capture import Region, ScreenCapture
from markitwrite.vision.analyzer import AnalysisResult, VisionAnalyzer


@dataclass
class StepLog:
    """One step in the assistant's execution pipeline."""

    step: str
    status: str  # "ok", "error", "skipped"
    detail: str = ""
    duration_ms: int = 0
    data: dict = field(default_factory=dict)


@dataclass
class AssistantResult:
    """Final result from the assistant."""

    success: bool
    output_path: Optional[str] = None
    summary: str = ""
    steps: list[StepLog] = field(default_factory=list)
    cropped_image: Optional[bytes] = None
    metadata: dict = field(default_factory=dict)


class VisionAssistant:
    """Multimodal file assistant: one sentence → screenshot → paste → done.

    Pipeline:
      1. Parse intent (Gemini 3 Flash via OpenRouter) → understand what user wants
      2. Take screenshot (mss/pyautogui) → full screen capture
      3. Analyze screenshot (Vision API) → find target region
      4. Crop region (PIL) → extract the relevant part
      5. Paste into document (markitwrite) → insert into DOCX/PPTX/MD
      6. (Optional) Verify result (Vision API) → QA check

    Example:
        assistant = VisionAssistant()

        # Full auto: screenshot → find → crop → paste
        result = assistant.run("把屏幕上的DCF模型截图放到report.docx里")

        # With an existing image (skip screenshot):
        result = assistant.run(
            "把这个表格放到slides.pptx第3页",
            image_path="table.png"
        )

        # Specify exact region (skip AI locate):
        result = assistant.run(
            "截图放到notes.md里",
            region=Region(left=100, top=200, width=800, height=600)
        )
    """

    def __init__(
        self,
        api_key: Optional[str] = None,
        model: Optional[str] = None,
        verbose: bool = True,
        on_step: Optional[Callable[[StepLog], None]] = None,
    ):
        """
        Args:
            api_key: OpenRouter API key. Uses OPENROUTER_API_KEY env var if None.
            model: Model to use. Defaults to google/gemini-3-flash-preview.
            verbose: Print progress to stderr.
            on_step: Callback for each pipeline step (for UI integration).
        """
        self._api_key = api_key
        self._model = model
        self._verbose = verbose
        self._on_step = on_step
        self._capture: Optional[ScreenCapture] = None
        self._analyzer: Optional[VisionAnalyzer] = None

    def _log(self, msg: str) -> None:
        if self._verbose:
            print(f"  [vision-assist] {msg}", file=sys.stderr)

    def _emit_step(self, step: StepLog) -> None:
        if self._on_step:
            self._on_step(step)

    @property
    def capture(self) -> ScreenCapture:
        if self._capture is None:
            self._capture = ScreenCapture()
        return self._capture

    @property
    def analyzer(self) -> VisionAnalyzer:
        if self._analyzer is None:
            self._analyzer = VisionAnalyzer(
                api_key=self._api_key, model=self._model
            )
        return self._analyzer

    def run(
        self,
        instruction: str,
        image_path: Optional[str] = None,
        image_bytes: Optional[bytes] = None,
        region: Optional[Region] = None,
        output_path: Optional[str] = None,
        monitor: int = 0,
        verify: bool = False,
    ) -> AssistantResult:
        """Execute the full pipeline from a single instruction.

        Args:
            instruction: Natural language command, e.g.
                "把屏幕上的DCF模型截图放到report.docx里"
            image_path: Use this image instead of taking a screenshot.
            image_bytes: Use these bytes instead of taking a screenshot.
            region: Skip AI locate, use this exact crop region.
            output_path: Override the output path (otherwise inferred from instruction).
            monitor: Which monitor to capture (0=all, 1=first, etc.)
            verify: If True, use Vision API to verify the result.

        Returns:
            AssistantResult with output path, summary, and step logs.
        """
        steps: list[StepLog] = []
        t_start = time.monotonic()

        # ── Step 1: Parse intent ──
        self._log("Step 1/5: Parsing intent...")
        t0 = time.monotonic()
        plan = self.analyzer.plan_from_text(instruction)
        step1 = StepLog(
            step="parse_intent",
            status="ok",
            detail=plan.get("reasoning", ""),
            duration_ms=int((time.monotonic() - t0) * 1000),
            data=plan,
        )
        steps.append(step1)
        self._emit_step(step1)
        self._log(f"  Intent: {plan.get('reasoning', '?')}")

        # Determine output path
        target_doc = output_path or plan.get("target_document", "output.docx")
        target_element = plan.get("target_element")
        needs_screenshot = plan.get("needs_screenshot", True)
        position = plan.get("position")
        size = plan.get("size")

        # ── Step 2: Get image ──
        screenshot_bytes: Optional[bytes] = None
        full_image: Optional[Image.Image] = None

        if image_bytes:
            self._log("Step 2/5: Using provided image bytes.")
            screenshot_bytes = image_bytes
            full_image = Image.open(__import__("io").BytesIO(image_bytes))
            steps.append(StepLog(step="get_image", status="ok", detail="from bytes"))
        elif image_path:
            self._log(f"Step 2/5: Loading image from {image_path}")
            with open(image_path, "rb") as f:
                screenshot_bytes = f.read()
            full_image = Image.open(image_path)
            steps.append(StepLog(step="get_image", status="ok", detail=f"from {image_path}"))
        elif needs_screenshot:
            self._log("Step 2/5: Taking screenshot...")
            t0 = time.monotonic()
            full_image = self.capture.take_screenshot(monitor=monitor)
            screenshot_bytes = self.capture.image_to_bytes(full_image)
            step2 = StepLog(
                step="screenshot",
                status="ok",
                detail=f"{full_image.width}x{full_image.height}",
                duration_ms=int((time.monotonic() - t0) * 1000),
            )
            steps.append(step2)
            self._emit_step(step2)
            self._log(f"  Captured: {full_image.width}x{full_image.height}")
        else:
            return AssistantResult(
                success=False,
                summary="No image source and screenshot not needed.",
                steps=steps,
            )

        # ── Step 3: Locate target region ──
        crop_region: Optional[Region] = region
        cropped_bytes: bytes = screenshot_bytes  # default: use full image

        if crop_region:
            self._log("Step 3/5: Using provided region, skipping AI locate.")
            steps.append(StepLog(step="locate", status="skipped", detail="manual region"))
        elif target_element and full_image:
            self._log(f"Step 3/5: Locating '{target_element}' on screen...")
            t0 = time.monotonic()
            locate_result = self.analyzer.locate(
                image_bytes=screenshot_bytes,
                target=target_element,
                image_width=full_image.width,
                image_height=full_image.height,
            )
            best = locate_result.best_region()
            duration = int((time.monotonic() - t0) * 1000)

            if best:
                crop_region = best
                self._log(
                    f"  Found: ({best.left},{best.top}) "
                    f"{best.width}x{best.height}"
                )
                steps.append(StepLog(
                    step="locate",
                    status="ok",
                    detail=f"Found at ({best.left},{best.top})",
                    duration_ms=duration,
                    data={"region": locate_result.regions},
                ))
            else:
                self._log("  Not found - using full screenshot.")
                steps.append(StepLog(
                    step="locate",
                    status="ok",
                    detail="Element not found, using full image",
                    duration_ms=duration,
                ))
        else:
            self._log("Step 3/5: No specific target, using full screenshot.")
            steps.append(StepLog(step="locate", status="skipped", detail="no target"))

        # ── Step 4: Crop ──
        if crop_region and full_image:
            self._log(f"Step 4/5: Cropping to {crop_region.width}x{crop_region.height}...")
            # Clamp region to image bounds
            clamped = Region(
                left=max(0, crop_region.left),
                top=max(0, crop_region.top),
                width=min(crop_region.width, full_image.width - max(0, crop_region.left)),
                height=min(crop_region.height, full_image.height - max(0, crop_region.top)),
            )
            cropped = self.capture.crop_image(full_image, clamped)
            cropped_bytes = self.capture.image_to_bytes(cropped)
            steps.append(StepLog(
                step="crop",
                status="ok",
                detail=f"{clamped.width}x{clamped.height}",
            ))
        else:
            self._log("Step 4/5: No crop needed, using full image.")
            steps.append(StepLog(step="crop", status="skipped"))

        # ── Step 5: Paste into document ──
        self._log(f"Step 5/5: Pasting into {target_doc}...")
        t0 = time.monotonic()
        try:
            from markitwrite import MarkItWrite

            writer = MarkItWrite()
            result = writer.paste(
                image_source=cropped_bytes,
                target=target_doc,
                position=position,
                size=size,
            )
            duration = int((time.monotonic() - t0) * 1000)
            steps.append(StepLog(
                step="paste",
                status="ok",
                detail=f"{len(result.output)} bytes → {target_doc}",
                duration_ms=duration,
            ))
            self._log(f"  Done: {len(result.output)} bytes → {target_doc}")
        except Exception as e:
            steps.append(StepLog(step="paste", status="error", detail=str(e)))
            return AssistantResult(
                success=False,
                output_path=target_doc,
                summary=f"Failed to paste: {e}",
                steps=steps,
                cropped_image=cropped_bytes,
            )

        # ── Optional: Verify ──
        if verify:
            self._log("Bonus: Verifying result with Vision API...")
            # We'd need to render the output doc as an image to verify.
            # For now, log as skipped - full verification needs doc-to-image.
            steps.append(StepLog(step="verify", status="skipped", detail="not yet implemented"))

        total_ms = int((time.monotonic() - t_start) * 1000)
        summary = (
            f"Captured '{target_element or 'screen'}' → "
            f"pasted into {target_doc} "
            f"({total_ms}ms total)"
        )
        self._log(f"All done! {summary}")

        return AssistantResult(
            success=True,
            output_path=target_doc,
            summary=summary,
            steps=steps,
            cropped_image=cropped_bytes,
            metadata={"total_ms": total_ms, "plan": plan},
        )

    def quick_capture(
        self,
        output_path: str = "output.docx",
        monitor: int = 0,
        size: Optional[dict] = None,
    ) -> AssistantResult:
        """Shortcut: take full screenshot and paste into a document. No AI needed."""
        self._log("Quick capture: screenshot → paste (no AI).")
        t0 = time.monotonic()
        img = self.capture.take_screenshot(monitor=monitor)
        img_bytes = self.capture.image_to_bytes(img)

        from markitwrite import MarkItWrite

        writer = MarkItWrite()
        result = writer.paste(
            image_source=img_bytes,
            target=output_path,
            size=size,
        )

        total_ms = int((time.monotonic() - t0) * 1000)
        return AssistantResult(
            success=True,
            output_path=output_path,
            summary=f"Full screenshot → {output_path} ({total_ms}ms)",
            steps=[
                StepLog(step="screenshot", status="ok", detail=f"{img.width}x{img.height}"),
                StepLog(step="paste", status="ok", detail=f"→ {output_path}"),
            ],
            cropped_image=img_bytes,
        )

    def capture_region(
        self,
        instruction: str,
        monitor: int = 0,
    ) -> tuple[bytes, Region | None]:
        """Just capture + locate, return the cropped image bytes.

        Useful when you want the image but will handle pasting yourself.
        """
        img = self.capture.take_screenshot(monitor=monitor)
        img_bytes = self.capture.image_to_bytes(img)

        plan = self.analyzer.plan_from_text(instruction)
        target = plan.get("target_element")

        if not target:
            return img_bytes, None

        locate_result = self.analyzer.locate(
            image_bytes=img_bytes,
            target=target,
            image_width=img.width,
            image_height=img.height,
        )
        best = locate_result.best_region()

        if best:
            cropped = self.capture.crop_image(img, best)
            return self.capture.image_to_bytes(cropped), best
        return img_bytes, None
