#!/usr/bin/env python3
"""
Helper script for syncing Antigravity brain markdown files with .agent-memory/
Used by sync_push.bat and sync_pull.bat in the workspace root.
"""

import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

WORKSPACE_ROOT = Path(__file__).resolve().parent
MEMORY_DIR = WORKSPACE_ROOT / ".agent-memory"
BRAIN_DIR = Path(os.path.expandvars(r"%USERPROFILE%\.gemini\antigravity\brain"))

ARTIFACT_MAP = {
    "task.md": "TASKS.md",
    "implementation_plan.md": "IMPLEMENTATION_PLAN.md",
    "walkthrough.md": "WALKTHROUGH.md",
}


def sync_brain_to_memory() -> None:
    MEMORY_DIR.mkdir(parents=True, exist_ok=True)
    if not BRAIN_DIR.exists():
        print(f"[INFO] Brain directory not found at {BRAIN_DIR}. Skipping brain artifact scan.")
        return

    print(f"[1/2] Scanning latest session artifacts in: {BRAIN_DIR}")
    for brain_name, memory_upper_name in ARTIFACT_MAP.items():
        candidates = list(BRAIN_DIR.glob(f"*/{brain_name}"))
        if not candidates:
            continue

        newest_src = max(candidates, key=lambda p: p.stat().st_mtime)
        src_mtime = newest_src.stat().st_mtime
        src_dt = datetime.fromtimestamp(src_mtime).strftime("%Y-%m-%d %H:%M:%S")

        upper_dst = MEMORY_DIR / memory_upper_name

        # Overwrite the canonical file in .agent-memory/ if the brain file is newer
        if not upper_dst.exists() or src_mtime > upper_dst.stat().st_mtime:
            shutil.copy2(newest_src, upper_dst)
            print(f"  -> Synced {brain_name} ({src_dt}) -> .agent-memory/{memory_upper_name}")
        else:
            dst_dt = datetime.fromtimestamp(upper_dst.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            print(f"  -> .agent-memory/{memory_upper_name} is already up to date ({dst_dt})")


def show_memory_status() -> None:
    print("\n=== Current .agent-memory/ Status ===")
    if not MEMORY_DIR.exists():
        print("  [!] .agent-memory/ directory does not exist yet.")
        return

    for md_file in sorted(MEMORY_DIR.glob("*.md")):
        mtime = datetime.fromtimestamp(md_file.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        size_kb = md_file.stat().st_size / 1024.0
        print(f"  - {md_file.name:<26} | Updated: {mtime} | {size_kb:6.1f} KB")
    print("=====================================\n")


def main() -> None:
    mode = sys.argv[1].lower() if len(sys.argv) > 1 else "status"
    if mode == "push":
        sync_brain_to_memory()
        show_memory_status()
    elif mode == "pull":
        show_memory_status()
    else:
        show_memory_status()


if __name__ == "__main__":
    main()
