"""Screen capture module - the 'eyes' of the assistant.

Uses `mss` for fast, cross-platform screen capture.
Falls back to `pyautogui` if mss is unavailable.
Provides region-based cropping via PIL.
"""

from __future__ import annotations

import io
import time
from dataclasses import dataclass
from typing import Optional

from PIL import Image


@dataclass
class Region:
    """Screen region defined by pixel coordinates."""

    left: int
    top: int
    width: int
    height: int

    @property
    def right(self) -> int:
        return self.left + self.width

    @property
    def bottom(self) -> int:
        return self.top + self.height

    @property
    def bbox(self) -> tuple[int, int, int, int]:
        """PIL-style (left, top, right, bottom) bounding box."""
        return (self.left, self.top, self.right, self.bottom)


class ScreenCapture:
    """Cross-platform screen capture with optional region cropping."""

    def __init__(self):
        self._backend = self._detect_backend()

    @staticmethod
    def _detect_backend() -> str:
        try:
            import mss  # noqa: F401
            return "mss"
        except ImportError:
            pass
        try:
            import pyautogui  # noqa: F401
            return "pyautogui"
        except ImportError:
            pass
        raise ImportError(
            "Screen capture requires 'mss' or 'pyautogui'. "
            "Install with: pip install mss  (recommended)"
        )

    def take_screenshot(
        self,
        monitor: int = 0,
        region: Optional[Region] = None,
    ) -> Image.Image:
        """Capture the screen (or a region of it).

        Args:
            monitor: Monitor index (0 = all monitors combined, 1 = first, etc.)
            region: Optional region to crop. If None, captures full screen.

        Returns:
            PIL Image of the capture.
        """
        if self._backend == "mss":
            img = self._capture_mss(monitor)
        else:
            img = self._capture_pyautogui()

        if region:
            img = img.crop(region.bbox)

        return img

    def take_screenshot_bytes(
        self,
        monitor: int = 0,
        region: Optional[Region] = None,
        fmt: str = "PNG",
    ) -> bytes:
        """Capture screen and return as bytes (ready for markitwrite or API)."""
        img = self.take_screenshot(monitor=monitor, region=region)
        buf = io.BytesIO()
        img.save(buf, format=fmt)
        return buf.getvalue()

    def _capture_mss(self, monitor: int) -> Image.Image:
        import mss

        with mss.mss() as sct:
            mon = sct.monitors[monitor]
            shot = sct.grab(mon)
            return Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")

    def _capture_pyautogui(self) -> Image.Image:
        import pyautogui

        return pyautogui.screenshot()

    @staticmethod
    def crop_image(image: Image.Image, region: Region) -> Image.Image:
        """Crop an existing PIL image to the given region."""
        return image.crop(region.bbox)

    @staticmethod
    def save_image(image: Image.Image, path: str, fmt: str = "PNG") -> str:
        """Save PIL image to disk, return the path."""
        image.save(path, format=fmt)
        return path

    @staticmethod
    def image_to_bytes(image: Image.Image, fmt: str = "PNG") -> bytes:
        """Convert PIL image to bytes."""
        buf = io.BytesIO()
        image.save(buf, format=fmt)
        return buf.getvalue()
