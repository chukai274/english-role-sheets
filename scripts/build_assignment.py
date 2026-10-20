#!/usr/bin/env python3
"""Generate and verify all role-sheet deliverables from one JSON file."""

from __future__ import annotations

import argparse
import copy
import json
import subprocess
import sys
from pathlib import Path

from build_role_docs import DEFAULT_TEMPLATE as DEFAULT_DOCX_TEMPLATE
from build_role_docs import SCHEMA as ROLE_SCHEMA
from build_role_docs import build as build_role_docs
from build_study_guide import DEFAULT_TEMPLATE as DEFAULT_GUIDE_TEMPLATE
from build_study_guide import SCHEMA as STUDY_GUIDE_SCHEMA
from build_study_guide import build as build_study_guide
from verify_outputs import verify


def combined_schema() -> dict:
    schema = copy.deepcopy(ROLE_SCHEMA)
    schema["study_guide"] = copy.deepcopy(STUDY_GUIDE_SCHEMA)
    return schema


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--content", type=Path, help="Combined UTF-8 assignment JSON")
    parser.add_argument("--output-dir", type=Path, help="Directory for all six deliverables")
    parser.add_argument("--docx-template", type=Path, default=DEFAULT_DOCX_TEMPLATE)
    parser.add_argument("--guide-template", type=Path, default=DEFAULT_GUIDE_TEMPLATE)
    parser.add_argument("--html-only", action="store_true", help="Skip PDF export")
    parser.add_argument("--verify-only", action="store_true", help="Verify existing outputs without rebuilding")
    parser.add_argument("--print-schema", action="store_true", help="Print the combined JSON schema")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.print_schema:
        print(json.dumps(combined_schema(), ensure_ascii=False, indent=2))
        return 0
    if args.content is None or args.output_dir is None:
        print("error: --content and --output-dir are required unless --print-schema is used", file=sys.stderr)
        return 2

    content_path = args.content.resolve()
    output_dir = args.output_dir.resolve()
    try:
        outputs: list[Path] = []
        if not args.verify_only:
            outputs.extend(
                build_role_docs(content_path, output_dir, args.docx_template.resolve())
            )
            outputs.extend(
                build_study_guide(
                    content_path,
                    output_dir,
                    args.guide_template.resolve(),
                    html_only=args.html_only,
                )
            )
        messages = verify(content_path, output_dir, require_pdf=not args.html_only)
    except (OSError, ValueError, json.JSONDecodeError, subprocess.SubprocessError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1

    for output in outputs:
        print(output)
    for message in messages:
        print(message)
    print("All deliverables generated and structurally verified.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
