#!/usr/bin/env python3
"""Run structural checks on all English role-sheet deliverables."""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path

from docx import Document
from pypdf import PdfReader

from build_role_docs import (
    ANALYZER_STRUCTURES,
    LEADER_PARTICIPATION,
    PURPOSES,
    ROLE_SPECS,
    SUMMARY_STRUCTURES,
    WORD_CATEGORIES,
    role_display_name,
    safe_filename,
)
from build_study_guide import IMAGE_REFERENCE_RE, metadata_for, output_basename


def fail(message: str) -> None:
    raise ValueError(message)


def check(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def document_text(document: Document) -> str:
    paragraphs = [paragraph.text for paragraph in document.paragraphs]
    cells = [
        cell.text
        for table in document.tables
        for row in table.rows
        for cell in row.cells
    ]
    return "\n".join([*paragraphs, *cells])


def check_docx_archive(path: Path) -> None:
    with zipfile.ZipFile(path) as archive:
        bad_member = archive.testzip()
        check(bad_member is None, f"Corrupt DOCX member in {path.name}: {bad_member}")
        check("word/document.xml" in archive.namelist(), f"Missing word/document.xml in {path.name}")


def expected_role_path(output_dir: Path, role_index: int, key: str, label: str, roles: dict) -> Path:
    owner = safe_filename(str(roles[key]["name"]))
    role_name = label.replace(" ", "_")
    return output_dir / f"{role_index}_{owner}_{role_name}.docx"


def verify_docx_files(data: dict, output_dir: Path) -> list[str]:
    roles = data["roles"]
    messages: list[str] = []

    documents: dict[str, Document] = {}
    texts: dict[str, str] = {}
    for role_index, (key, label, marker) in enumerate(ROLE_SPECS, 1):
        path = expected_role_path(output_dir, role_index, key, label, roles)
        check(path.is_file(), f"Missing role document: {path}")
        check_docx_archive(path)
        document = Document(path)
        text = document_text(document)
        documents[key] = document
        texts[key] = text

        check(marker in text, f"{path.name} does not contain its assigned role heading")
        other_markers = [other for _, _, other in ROLE_SPECS if other != marker]
        check(not any(other in text for other in other_markers), f"{path.name} contains another role sheet")
        check(str(data["title"]) in text, f"{path.name} is missing the article title")
        visible_owner = role_display_name(roles[key])
        check(visible_owner in text, f"{path.name} is missing the owner name")
        check(not IMAGE_REFERENCE_RE.search(text), f"{path.name} contains an image-filename reference")
        check(not re.search(r"_{4,}", text), f"{path.name} still contains a template blank line")
        messages.append(f"OK DOCX: {path.name}")

    leader = roles["discussion_leader"]
    leader_doc = documents["discussion_leader"]
    leader_text = texts["discussion_leader"]
    selected = leader.get("participation")
    if selected is None:
        selected = [key for key, _ in LEADER_PARTICIPATION] if leader.get("meeting_completed") else []
    expected_checked = len(selected)
    check(leader_text.count("☑") == expected_checked, "Discussion Leader checked-box count is wrong")
    check(
        leader_text.count("☐") == len(LEADER_PARTICIPATION) - expected_checked,
        "Discussion Leader unchecked-box count is wrong",
    )
    check(len(leader_doc.tables) == 2, "Discussion Leader answer table is missing")
    answer_rows = leader_doc.tables[1].rows[1:]
    check(len(answer_rows) == 4, "Discussion Leader must contain four answer rows")
    if leader.get("meeting_completed"):
        check(
            all(row.cells[1].text.strip() and row.cells[2].text.strip() for row in answer_rows),
            "Completed Discussion Leader sheet has blank answerers or best ideas",
        )

    summarizer = roles["summarizer"]
    summarizer_text = texts["summarizer"]
    summarizer_doc = documents["summarizer"]
    check(len(summarizer_doc.tables) == 2, "Summarizer details chart table is missing")
    check(len(summarizer_doc.tables[1].rows) == 5, "Summarizer details chart must contain four section rows")
    check(
        all(all(cell.text.strip() for cell in row.cells) for row in summarizer_doc.tables[1].rows),
        "Summarizer details chart contains a blank field",
    )
    selected_summary = len(summarizer["purposes"]) + len(summarizer["structures"])
    summary_total = len(PURPOSES) + len(SUMMARY_STRUCTURES)
    check(summarizer_text.count("☑") == selected_summary, "Summarizer checked-box count is wrong")
    check(summarizer_text.count("☐") == summary_total - selected_summary, "Summarizer unchecked-box count is wrong")
    short_words = re.findall(r"\b[\w'-]+\b", str(summarizer["short_summary"]))
    check(len(short_words) <= 40, "Summarizer short summary exceeds 40 words")

    word_master = roles["word_master"]
    word_text = texts["word_master"]
    check(word_text.count("☑") == len(word_master["categories"]), "Word Master checked-box count is wrong")
    check(
        word_text.count("☐") == len(WORD_CATEGORIES) - len(word_master["categories"]),
        "Word Master unchecked-box count is wrong",
    )
    word_doc = documents["word_master"]
    check(len(word_doc.tables) == 2, "Word Master vocabulary table is missing")
    check(
        len(word_doc.tables[1].rows) == len(word_master["words"]) + 1,
        "Word Master vocabulary row count is wrong",
    )
    check(
        all(all(cell.text.strip() for cell in row.cells) for row in word_doc.tables[1].rows[1:]),
        "Word Master contains a blank vocabulary field",
    )

    analyzer = roles["structure_analyzer"]
    analyzer_text = texts["structure_analyzer"]
    check(analyzer_text.count("☑") == len(analyzer["structures"]), "Structure Analyzer checked-box count is wrong")
    check(
        analyzer_text.count("☐") == len(ANALYZER_STRUCTURES) - len(analyzer["structures"]),
        "Structure Analyzer unchecked-box count is wrong",
    )
    analyzer_doc = documents["structure_analyzer"]
    check(len(analyzer_doc.tables) == 2, "Structure Analyzer map table is missing")
    check(len(analyzer_doc.tables[1].rows) == 5, "Structure Analyzer map must contain four rows")
    check(
        all(all(cell.text.strip() for cell in row.cells) for row in analyzer_doc.tables[1].rows[1:]),
        "Structure Analyzer contains a blank map field",
    )
    return messages


def verify_study_guide(data: dict, output_dir: Path, *, require_pdf: bool) -> list[str]:
    basename = output_basename(data)
    html_path = output_dir / f"{basename}.html"
    pdf_path = output_dir / f"{basename}.pdf"
    messages: list[str] = []

    check(html_path.is_file(), f"Missing study-guide HTML: {html_path}")
    html_text = html_path.read_text(encoding="utf-8")
    check("{{" not in html_text and "}}" not in html_text, "Study-guide HTML contains template markers")
    check(not IMAGE_REFERENCE_RE.search(html_text), "Study-guide HTML contains an image-filename reference")
    check(str(data["study_guide"]["chinese_title"]) in html_text, "Study-guide HTML is missing its title")
    check("教材第" in html_text, "Study-guide HTML is missing printed page references")
    messages.append(f"OK HTML: {html_path.name}")

    if require_pdf:
        check(pdf_path.is_file(), f"Missing study-guide PDF: {pdf_path}")
        reader = PdfReader(pdf_path)
        check(len(reader.pages) > 0, "Study-guide PDF has no pages")
        text = "\n".join((page.extract_text() or "") for page in reader.pages)
        check(not IMAGE_REFERENCE_RE.search(text), "Study-guide PDF contains an image-filename reference")
        check("教材第" in text, "Study-guide PDF is missing printed page references")
        expected_meta = metadata_for(data)
        metadata = reader.metadata
        check(metadata.title == expected_meta["title"], "Study-guide PDF title metadata is wrong")
        check(metadata.author == expected_meta["author"], "Study-guide PDF author metadata is wrong")
        check(bool(metadata.subject), "Study-guide PDF subject metadata is missing")
        check(bool(metadata.get("/Keywords")), "Study-guide PDF keyword metadata is missing")
        page_note = ""
        if not 5 <= len(reader.pages) <= 8:
            page_note = f" (review page density: {len(reader.pages)} pages)"
        messages.append(f"OK PDF: {pdf_path.name}{page_note}")
    return messages


def verify(content_path: Path, output_dir: Path, *, require_pdf: bool = True) -> list[str]:
    data = json.loads(content_path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        fail("Content JSON root must be an object")
    messages = verify_docx_files(data, output_dir)
    messages.extend(verify_study_guide(data, output_dir, require_pdf=require_pdf))
    return messages


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--content", type=Path, required=True, help="Combined assignment JSON")
    parser.add_argument("--output-dir", type=Path, required=True, help="Directory containing generated files")
    parser.add_argument("--html-only", action="store_true", help="Do not require a PDF")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        messages = verify(args.content.resolve(), args.output_dir.resolve(), require_pdf=not args.html_only)
    except (OSError, ValueError, json.JSONDecodeError, zipfile.BadZipFile) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    for message in messages:
        print(message)
    print("All structural checks passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
