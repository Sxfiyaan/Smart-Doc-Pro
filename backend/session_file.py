from dataclasses import dataclass
from pathlib import Path


@dataclass
class SessionFile:
    file_id: str
    display_name: str
    path: Path
    size: int
    is_output: bool = False
