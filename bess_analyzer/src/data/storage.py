"""Project save/load functionality using JSON serialization."""

import json
from pathlib import Path

from src.models.project import Project


def save_project(project: Project, filepath: str) -> None:
    """Save a project to a JSON file.

    Args:
        project: Project object to save.
        filepath: Output file path (should end in .json).

    Raises:
        OSError: If file cannot be written.
    """
    data = project.to_dict()
    path = Path(filepath)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, default=str)


def load_project(filepath: str) -> Project:
    """Load a project from a JSON file.

    Args:
        filepath: Path to the JSON project file.

    Returns:
        Reconstructed Project object.

    Raises:
        FileNotFoundError: If file does not exist.
        json.JSONDecodeError: If file is not valid JSON.
        KeyError: If required fields are missing.
    """
    with open(filepath, "r", encoding="utf-8") as f:
        data = json.load(f)
    return Project.from_dict(data)
