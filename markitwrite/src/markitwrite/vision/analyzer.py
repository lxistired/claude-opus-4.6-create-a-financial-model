"""Claude Vision analyzer - the 'brain' that understands what's on screen.

Sends screenshots to Claude's multimodal API and gets back:
- Content description (what's in the image)
- Region coordinates (where a specific element is)
- Structured extraction (tables, charts, text)
"""

from __future__ import annotations

import base64
import json
import re
from dataclasses import dataclass, field
from typing import Any, Optional

from markitwrite.vision.capture import Region


@dataclass
class AnalysisResult:
    """Structured result from Claude Vision analysis."""

    description: str = ""
    regions: list[dict] = field(default_factory=list)
    raw_response: str = ""
    metadata: dict = field(default_factory=dict)

    def best_region(self) -> Optional[Region]:
        """Return the highest-confidence detected region, or None."""
        if not self.regions:
            return None
        best = max(self.regions, key=lambda r: r.get("confidence", 0))
        return Region(
            left=best["left"],
            top=best["top"],
            width=best["width"],
            height=best["height"],
        )


class VisionAnalyzer:
    """Analyze images using Claude's multimodal Vision API.

    Requires: pip install anthropic
    Set ANTHROPIC_API_KEY env var or pass api_key to constructor.
    """

    DEFAULT_MODEL = "claude-sonnet-4-5-20250929"

    def __init__(
        self,
        api_key: Optional[str] = None,
        model: Optional[str] = None,
    ):
        import anthropic

        self._client = anthropic.Anthropic(api_key=api_key) if api_key else anthropic.Anthropic()
        self._model = model or self.DEFAULT_MODEL

    def describe(self, image_bytes: bytes, question: str = "") -> AnalysisResult:
        """Describe what's in an image.

        Args:
            image_bytes: PNG/JPEG image data.
            question: Optional question about the image (e.g. "这是什么表格？")

        Returns:
            AnalysisResult with description.
        """
        prompt = question or "Describe what you see in this image in detail."

        response = self._call_vision(image_bytes, prompt)
        return AnalysisResult(description=response, raw_response=response)

    def locate(
        self,
        image_bytes: bytes,
        target: str,
        image_width: int,
        image_height: int,
    ) -> AnalysisResult:
        """Find a specific element in the image and return its coordinates.

        Args:
            image_bytes: Full screenshot as PNG/JPEG.
            target: What to find (e.g. "DCF模型表格", "revenue chart").
            image_width: Actual pixel width of the image.
            image_height: Actual pixel height of the image.

        Returns:
            AnalysisResult with regions containing pixel-coordinate bounding boxes.
        """
        prompt = f"""I need you to locate a specific element in this screenshot.

Target to find: "{target}"

The image dimensions are {image_width}x{image_height} pixels.

Return a JSON object with the following structure:
{{
  "found": true/false,
  "description": "brief description of what you found",
  "regions": [
    {{
      "label": "name of the element",
      "left": <pixel x of top-left corner>,
      "top": <pixel y of top-left corner>,
      "width": <pixel width>,
      "height": <pixel height>,
      "confidence": <0.0 to 1.0>
    }}
  ]
}}

IMPORTANT:
- Coordinates must be in absolute pixels based on the {image_width}x{image_height} image.
- Add generous padding (50-100px) around the element so it looks good when cropped.
- If you find multiple matching elements, return all of them sorted by confidence.
- Return ONLY the JSON, no other text."""

        raw = self._call_vision(image_bytes, prompt)
        return self._parse_locate_response(raw)

    def decide_action(
        self,
        image_bytes: bytes,
        user_instruction: str,
    ) -> dict[str, Any]:
        """Given a screenshot and a natural language instruction, decide what to do.

        This is the 'brain' - it understands the user's intent and plans actions.

        Args:
            image_bytes: Current screenshot.
            user_instruction: e.g. "把DCF模型截图放到report.docx第3段后面"

        Returns:
            Action plan dict, e.g.:
            {
                "action": "screenshot_and_paste",
                "target_element": "DCF model table",
                "target_document": "report.docx",
                "position": {"paragraph": 3},
                "size": {"width": 6.0},
                "reasoning": "User wants the DCF valuation table..."
            }
        """
        prompt = f"""You are a smart file assistant. The user gave you this instruction:

"{user_instruction}"

Look at the screenshot and decide what action to take.

Return a JSON object:
{{
  "action": "screenshot_and_paste",
  "target_element": "<what to capture from the screen - describe it specifically>",
  "target_document": "<output file path, e.g. report.docx>",
  "position": {{"paragraph": N}} or {{"slide": N}} or {{"after_heading": "text"}} or null,
  "size": {{"width": 6.0}} or null,
  "reasoning": "<brief explanation of your understanding>"
}}

If the instruction doesn't require a screenshot (e.g. just text), set action to "text_only".
If you cannot determine the target document from the instruction, set it to "output.docx".
Return ONLY the JSON, no other text."""

        raw = self._call_vision(image_bytes, prompt)
        return self._parse_json_response(raw)

    def plan_from_text(self, user_instruction: str) -> dict[str, Any]:
        """Plan actions from text instruction alone (no screenshot yet).

        Used as the first step - understand what the user wants before
        deciding whether we even need a screenshot.
        """
        import anthropic

        response = self._client.messages.create(
            model=self._model,
            max_tokens=1024,
            messages=[
                {
                    "role": "user",
                    "content": f"""You are a smart file assistant that can:
1. Take screenshots of the screen
2. Analyze and crop specific regions from screenshots
3. Paste images into documents (DOCX, PPTX, Markdown)

The user said: "{user_instruction}"

Decide what to do. Return a JSON object:
{{
  "needs_screenshot": true/false,
  "target_element": "<what to look for on screen, or null>",
  "target_document": "<output file path>",
  "target_format": "<docx/pptx/md>",
  "position": {{"paragraph": N}} or {{"slide": N}} or {{"after_heading": "text"}} or null,
  "size": {{"width": 6.0}} or null,
  "reasoning": "<your understanding of the task>"
}}

Return ONLY the JSON, no other text.""",
                }
            ],
        )
        raw = response.content[0].text
        return self._parse_json_response(raw)

    def verify_result(
        self,
        before_bytes: bytes,
        after_bytes: bytes,
        instruction: str,
    ) -> dict[str, Any]:
        """Verify the result by comparing before/after (optional QA step)."""
        prompt = f"""The user asked: "{instruction}"

I'm showing you two images:
1. First image: the original screenshot (what was captured)
2. Second image: the resulting document (after pasting)

Verify:
- Was the correct content captured?
- Was it placed correctly in the document?

Return JSON:
{{
  "success": true/false,
  "issues": ["list of any problems found"],
  "summary": "brief description of what was done"
}}

Return ONLY the JSON."""

        b64_before = base64.standard_b64encode(before_bytes).decode("utf-8")
        b64_after = base64.standard_b64encode(after_bytes).decode("utf-8")

        import anthropic

        response = self._client.messages.create(
            model=self._model,
            max_tokens=1024,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": prompt,
                        },
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": "image/png",
                                "data": b64_before,
                            },
                        },
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": "image/png",
                                "data": b64_after,
                            },
                        },
                    ],
                }
            ],
        )
        raw = response.content[0].text
        return self._parse_json_response(raw)

    # ── internals ──

    def _call_vision(self, image_bytes: bytes, prompt: str) -> str:
        """Send image + prompt to Claude Vision, return text response."""
        b64 = base64.standard_b64encode(image_bytes).decode("utf-8")

        # Detect media type from magic bytes
        media_type = "image/png"
        if image_bytes[:2] == b"\xff\xd8":
            media_type = "image/jpeg"
        elif image_bytes[:4] == b"RIFF" and image_bytes[8:12] == b"WEBP":
            media_type = "image/webp"

        response = self._client.messages.create(
            model=self._model,
            max_tokens=2048,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": b64,
                            },
                        },
                        {
                            "type": "text",
                            "text": prompt,
                        },
                    ],
                }
            ],
        )
        return response.content[0].text

    @staticmethod
    def _parse_json_response(raw: str) -> dict:
        """Extract JSON from Claude's response (handles markdown fences)."""
        # Try direct parse first
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            pass

        # Try extracting from ```json ... ``` blocks
        match = re.search(r"```(?:json)?\s*\n?(.*?)\n?\s*```", raw, re.DOTALL)
        if match:
            try:
                return json.loads(match.group(1))
            except json.JSONDecodeError:
                pass

        # Try finding first { ... } block
        match = re.search(r"\{.*\}", raw, re.DOTALL)
        if match:
            try:
                return json.loads(match.group(0))
            except json.JSONDecodeError:
                pass

        return {"error": "Failed to parse response", "raw": raw}

    @staticmethod
    def _parse_locate_response(raw: str) -> AnalysisResult:
        """Parse a locate() response into AnalysisResult."""
        data = VisionAnalyzer._parse_json_response(raw)

        if "error" in data:
            return AnalysisResult(
                description=data.get("raw", "Parse error"),
                raw_response=raw,
            )

        regions = []
        for r in data.get("regions", []):
            regions.append(
                {
                    "label": r.get("label", ""),
                    "left": int(r.get("left", 0)),
                    "top": int(r.get("top", 0)),
                    "width": int(r.get("width", 0)),
                    "height": int(r.get("height", 0)),
                    "confidence": float(r.get("confidence", 0.0)),
                }
            )

        return AnalysisResult(
            description=data.get("description", ""),
            regions=regions,
            raw_response=raw,
            metadata={"found": data.get("found", False)},
        )
