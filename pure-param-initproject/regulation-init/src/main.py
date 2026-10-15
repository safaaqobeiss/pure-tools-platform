"""
Entry point for the Regulation Project Initializer.

Usage:
    python main.py
"""
import sys
from pathlib import Path

# Ensure src/ is importable when running from repo root
sys.path.insert(0, str(Path(__file__).resolve().parent))


def main():
    from gui import run_gui
    run_gui()


if __name__ == "__main__":
    main()
