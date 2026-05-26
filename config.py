"""Shared path configuration — auto-detects Google Colab vs local machine."""

from pathlib import Path

COLAB_INPUT_DIR = "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team /Documents/AI Adoption RMT/RMT_APTIV_VERSIGENT/RMT_Air/input"
COLAB_OUTPUT_DIR = "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team /Documents/AI Adoption RMT/RMT_APTIV_VERSIGENT/RMT_Air/output"
COLAB_PROCESSING_DIR = "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team /Documents/AI Adoption RMT/RMT_APTIV_VERSIGENT/RMT_Air/processing"

def _is_colab() -> bool:
    try:
        import google.colab  # noqa: F401
        return True
    except ImportError:
        return False

IS_COLAB = _is_colab()

if IS_COLAB:
    INPUT_DIR = Path(COLAB_INPUT_DIR)
    OUTPUT_DIR = Path(COLAB_OUTPUT_DIR)
    PROCESSING_DIR = Path(COLAB_PROCESSING_DIR)
else:
    _BASE = Path(__file__).parent
    INPUT_DIR = _BASE / "input"
    OUTPUT_DIR = _BASE / "output"
    PROCESSING_DIR = _BASE / "processing"
