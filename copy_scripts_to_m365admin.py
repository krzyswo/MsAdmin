import os
import re
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DEST = ROOT / "M365Admin"

# Directories to skip while walking
SKIP_DIRS = {".git", "M365Admin", "__pycache__"}


def kebabify(name: str) -> str:
    base, ext = os.path.splitext(name)
    cleaned = re.sub(r"[^A-Za-z0-9]+", "-", base).strip("-")
    cleaned = re.sub(r"-+", "-", cleaned).lower()
    return f"{cleaned}{ext.lower()}" if ext else cleaned


def collision_safe(path: Path, seen: set[Path]) -> Path:
    if path not in seen and not path.exists():
        return path
    counter = 2
    stem, suffix = path.stem, path.suffix
    parent = path.parent
    while True:
        candidate = parent / f"{stem}-{counter}{suffix}"
        if candidate not in seen and not candidate.exists():
            return candidate
        counter += 1


def main() -> None:
    if DEST.exists():
        shutil.rmtree(DEST)
    DEST.mkdir()

    seen_paths: set[Path] = set()
    for root, dirs, files in os.walk(ROOT):
        # prune directories
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS and not d.startswith('.')]

        for filename in files:
            if not filename.lower().endswith(".ps1"):
                continue

            src_path = Path(root) / filename
            # Build mirrored path under DEST with kebab-case segments
            relative_parts = Path(root).relative_to(ROOT).parts
            kebab_parts = [kebabify(part) for part in relative_parts]
            kebab_file = kebabify(filename)
            dest_dir = DEST.joinpath(*kebab_parts)
            dest_dir.mkdir(parents=True, exist_ok=True)
            candidate = dest_dir / kebab_file
            candidate = collision_safe(candidate, seen_paths)
            shutil.copy2(src_path, candidate)
            seen_paths.add(candidate)


if __name__ == "__main__":
    main()
