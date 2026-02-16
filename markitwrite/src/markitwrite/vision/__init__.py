"""markitwrite.vision - Multimodal vision assistant for screen capture and document insertion.

Usage:
    from markitwrite.vision import VisionAssistant

    assistant = VisionAssistant()
    result = assistant.run("把屏幕上的DCF模型截图放到report.docx里")

Components:
    - ScreenCapture: Cross-platform screen capture (mss/pyautogui)
    - VisionAnalyzer: Claude Vision API for image understanding
    - VisionAssistant: Orchestrator that ties it all together
"""

from markitwrite.vision.capture import Region, ScreenCapture
from markitwrite.vision.analyzer import AnalysisResult, VisionAnalyzer
from markitwrite.vision.assistant import AssistantResult, VisionAssistant

__all__ = [
    "Region",
    "ScreenCapture",
    "AnalysisResult",
    "VisionAnalyzer",
    "AssistantResult",
    "VisionAssistant",
]
