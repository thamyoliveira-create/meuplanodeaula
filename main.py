# -*- coding: utf-8 -*-
"""Lesson Plan Generator - CLI entry point

Usage:
    python -m lesson_plan_generator.main --spreadsheet <path> --week <number> [--output <file>]

Arguments:
    --spreadsheet   Path to the Excel/CSV file containing the course schedule.
    --week          Week number (or identifier) to generate the lesson plan for.
    --output        Optional output PDF file path. Defaults to 'lesson_plan_week_<week>.pdf'.
"""

import argparse
import pathlib
import sys

from .plan_generator import generate_lesson_plan


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate a lesson plan PDF for a given week.")
    parser.add_argument(
        "--spreadsheet",
        required=True,
        type=pathlib.Path,
        help="Path to the Excel (.xlsx) or CSV file with the schedule.",
    )
    parser.add_argument(
        "--week",
        required=True,
        type=str,
        help="Week identifier (e.g., '1', '2', or 'Semana 1').",
    )
    parser.add_argument(
        "--output",
        type=pathlib.Path,
        default=None,
        help="Destination PDF file. If omitted, a file named 'lesson_plan_week_<week>.pdf' is created.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        output_path = args.output
        if output_path is None:
            output_path = pathlib.Path(f"lesson_plan_week_{args.week}.pdf")
        generate_lesson_plan(
            spreadsheet_path=args.spreadsheet,
            week=args.week,
            output_pdf_path=output_path,
        )
        print(f"✅ Lesson plan generated: {output_path}")
        return 0
    except Exception as exc:
        print(f"❌ Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
