"""Confirm a revenue SQT after PAD verifies the saved website value."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.revenue_pad import RevenuePadError, finish_selection, mark_sqt_imported


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("command", choices=("mark-imported", "finish"))
    parser.add_argument("--selection", required=True)
    parser.add_argument("--sqt")
    args = parser.parse_args()
    try:
        if args.command == "mark-imported":
            if not args.sqt:
                parser.error("--sqt is required for mark-imported")
            print(f"Marked {mark_sqt_imported(args.selection, args.sqt)} rows for SQT {args.sqt}")
        else:
            finish_selection(args.selection)
            print("Revenue batch complete")
    except RevenuePadError as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
