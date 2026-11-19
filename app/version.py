"""
Version and repository metadata for AssignmentFixer.
"""
from typing import Tuple

APP_VERSION = "1.0.4"
GITHUB_REPO = "tunkyothedev/NEU-AssignmentFixer"

def parse_version(ver_str: str) -> Tuple[int, ...]:
    """
    Parse a semantic version string (e.g., 'v1.2.3' or '1.2.0') into a tuple of ints.
    """
    cleaned = ver_str.strip().lstrip("vV")
    parts = []
    for seg in cleaned.split("."):
        try:
            parts.append(int(seg))
        except ValueError:
            break
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts)

def is_newer_version(latest_str: str, current_str: str = APP_VERSION) -> bool:
    """
    Returns True if latest_str is strictly newer than current_str.
    """
    return parse_version(latest_str) > parse_version(current_str)
