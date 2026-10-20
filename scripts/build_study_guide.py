#!/usr/bin/env python3
"""Build the Kami Chinese study guide HTML and PDF from assignment JSON."""

from __future__ import annotations

import argparse
import datetime
import html
import json
import os
import re
import shutil
import signal
import subprocess
import sys
import tempfile
import time
from pathlib import Path

from pypdf import PdfReader, PdfWriter


SKILL_DIR = Path(__file__).resolve().parent.parent
DEFAULT_TEMPLATE = SKILL_DIR / "assets" / "study-guide-template.html"

IMAGE_REFERENCE_RE = re.compile(
    r"(?:\bimage\s*\d+\b|图片\s*\d+|第\s*\d+\s*张图片)",
    re.IGNORECASE,
)

SCHEMA = {
    "week_label": "Week 3",
    "source_pages": "30-32",
    "paragraph_range": "P1-P18",
    "chinese_title": "Chinese study-guide title",
    "overview": "Concise Chinese overview",
    "central_claim": "Central claim or recommended action",
    "structure_parts": [
        {
            "page_range": "Pages 30-31",
            "paragraph_range": "P1-P4",
            "function": "问题与价值",
            "content": "Chinese explanation",
        }
    ],
    "reading_tips": ["Chinese reading tip"],
    "translations": [
        {
            "paragraph": "P1",
            "page": "Page 30",
            "text": "Chinese translation",
            "page_break_before": False,
        }
    ],
    "vocabulary": [
        {
            "word": "digital legacy",
            "pronunciation": "/.../",
            "part_of_speech": "n.",
            "meaning_zh": "Chinese meaning",
            "context": "Exact or shortened source context",
        }
    ],
    "star_word": {
        "word": "digital executor",
        "memory_picture_zh": "Chinese memory picture",
        "example": "Original English example sentence",
        "question_zh": "Chinese review question",
    },
    "additional_phrases": [
        {"phrase": "be left in limbo", "meaning_zh": "Chinese meaning"}
    ],
    "review_prompts": ["Chinese retelling prompt"],
    "discussion_questions": ["Four Chinese discussion questions"],
    "checklist": [
        {"text": "Chinese self-check item", "checked": False}
    ],
    "takeaway": "One-sentence Chinese takeaway",
    "metadata": {
        "author": "English Reading Circle",
        "description": "PDF description (optional)",
        "keywords": ["keyword 1", "keyword 2"],
    },
    "output_basename": "Optional output filename stem",
}


def fail(message: str) -> None:
    raise ValueError(message)


def required(mapping: dict, key: str, where: str) -> object:
    value = mapping.get(key)
    if value is None or value == "" or value == []:
        fail(f"Missing {where}.{key}")
    return value


def require_list(
    mapping: dict,
    key: str,
    where: str,
    *,
    exact: int | None = None,
    minimum: int | None = None,
    maximum: int | None = None,
) -> list:
    value = required(mapping, key, where)
    if not isinstance(value, list):
        fail(f"{where}.{key} must be a list")
    if exact is not None and len(value) != exact:
        fail(f"{where}.{key} must contain exactly {exact} items")
    if minimum is not None and len(value) < minimum:
        fail(f"{where}.{key} must contain at least {minimum} items")
    if maximum is not None and len(value) > maximum:
        fail(f"{where}.{key} must contain at most {maximum} items")
    return value


def iter_strings(value: object):
    if isinstance(value, str):
        yield value
    elif isinstance(value, dict):
        for child in value.values():
            yield from iter_strings(child)
    elif isinstance(value, list):
        for child in value:
            yield from iter_strings(child)


def validate(data: dict) -> None:
    required(data, "title", "root")
    roles = required(data, "roles", "root")
    if not isinstance(roles, dict):
        fail("root.roles must be an object")
    word_master = required(roles, "word_master", "roles")
    if not isinstance(word_master, dict):
        fail("roles.word_master must be an object")
    required(word_master, "name", "roles.word_master")

    guide = required(data, "study_guide", "root")
    if not isinstance(guide, dict):
        fail("root.study_guide must be an object")

    for field in (
        "week_label",
        "source_pages",
        "paragraph_range",
        "chinese_title",
        "overview",
        "central_claim",
        "takeaway",
    ):
        required(guide, field, "study_guide")

    if not re.search(r"\d", str(guide["source_pages"])):
        fail("study_guide.source_pages must contain printed page numbers")

    parts = require_list(guide, "structure_parts", "study_guide", exact=4)
    for index, part in enumerate(parts, 1):
        for field in ("page_range", "paragraph_range", "function", "content"):
            required(part, field, f"study_guide.structure_parts[{index}]")
        if not re.search(r"\d", str(part["page_range"])):
            fail(f"study_guide.structure_parts[{index}].page_range must contain printed page numbers")

    require_list(guide, "reading_tips", "study_guide", minimum=1)

    translations = require_list(guide, "translations", "study_guide", minimum=1)
    for index, item in enumerate(translations, 1):
        for field in ("paragraph", "page", "text"):
            required(item, field, f"study_guide.translations[{index}]")
        if not re.search(r"\d", str(item["page"])):
            fail(f"study_guide.translations[{index}].page must contain a printed page number")
    range_match = re.fullmatch(
        r"\s*P(\d+)\s*[-–—]\s*P(\d+)\s*",
        str(guide["paragraph_range"]),
        flags=re.IGNORECASE,
    )
    if range_match:
        start, end = map(int, range_match.groups())
        if end < start:
            fail("study_guide.paragraph_range must run from a lower paragraph number to a higher one")
        expected = [f"P{number}" for number in range(start, end + 1)]
        actual = [str(item["paragraph"]).strip().upper() for item in translations]
        if actual != expected:
            missing = [label for label in expected if label not in actual]
            duplicates = sorted({label for label in actual if actual.count(label) > 1})
            detail = []
            if missing:
                detail.append(f"missing {', '.join(missing)}")
            if duplicates:
                detail.append(f"duplicate {', '.join(duplicates)}")
            if not detail:
                detail.append("paragraphs are out of order")
            fail("study_guide.translations must cover the full paragraph range in order: " + "; ".join(detail))

    vocabulary = require_list(guide, "vocabulary", "study_guide", minimum=4, maximum=8)
    for index, item in enumerate(vocabulary, 1):
        for field in ("word", "pronunciation", "part_of_speech", "meaning_zh", "context"):
            required(item, field, f"study_guide.vocabulary[{index}]")
    role_words = {
        str(item.get("word", "")).strip().casefold()
        for item in word_master.get("words", [])
        if str(item.get("word", "")).strip()
    }
    guide_words = {str(item["word"]).strip().casefold() for item in vocabulary}
    missing_role_words = sorted(role_words - guide_words)
    if missing_role_words:
        fail("study_guide.vocabulary must include every Word Master entry: " + ", ".join(missing_role_words))

    star = required(guide, "star_word", "study_guide")
    if not isinstance(star, dict):
        fail("study_guide.star_word must be an object")
    for field in ("word", "memory_picture_zh", "example", "question_zh"):
        required(star, field, "study_guide.star_word")
    role_star = word_master.get("star") or {}
    if role_star.get("word") and str(star["word"]).strip().casefold() != str(role_star["word"]).strip().casefold():
        fail("study_guide.star_word.word must match roles.word_master.star.word")

    phrases = require_list(guide, "additional_phrases", "study_guide", minimum=1)
    for index, item in enumerate(phrases, 1):
        for field in ("phrase", "meaning_zh"):
            required(item, field, f"study_guide.additional_phrases[{index}]")

    require_list(guide, "review_prompts", "study_guide", minimum=3)
    require_list(guide, "discussion_questions", "study_guide", exact=4)
    checklist = require_list(guide, "checklist", "study_guide", minimum=3)
    for index, item in enumerate(checklist, 1):
        if isinstance(item, str):
            if not item.strip():
                fail(f"study_guide.checklist[{index}] must not be empty")
        elif isinstance(item, dict):
            required(item, "text", f"study_guide.checklist[{index}]")
        else:
            fail(f"study_guide.checklist[{index}] must be a string or object")

    for value in iter_strings(data):
        if IMAGE_REFERENCE_RE.search(value):
            fail(f"Use printed page numbers instead of image references: {value}")


def esc(value: object) -> str:
    return html.escape(str(value), quote=True)


def css_string(value: object) -> str:
    return str(value).replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ")


def clean_page_label(value: object) -> str:
    text = str(value).strip()
    text = re.sub(r"^(?:printed\s+)?pages?\s*:?[ ]*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"^第\s*", "", text)
    text = re.sub(r"\s*页$", "", text)
    return text


def safe_filename(value: str) -> str:
    cleaned = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "_", value).strip(" .")
    return cleaned or "study-guide"


def output_basename(data: dict) -> str:
    guide = data["study_guide"]
    configured = str(guide.get("output_basename", "")).strip()
    if configured:
        return safe_filename(configured)
    week = re.sub(r"\s+", "", str(guide["week_label"]))
    return safe_filename(f"{week}_中文学习资料")


def metadata_for(data: dict) -> dict[str, str]:
    guide = data["study_guide"]
    configured = guide.get("metadata") or {}
    if not isinstance(configured, dict):
        fail("study_guide.metadata must be an object")
    title = f"{guide['chinese_title']} 中文学习资料"
    author = str(configured.get("author") or "English Reading Circle")
    description = str(configured.get("description") or guide["overview"])
    keywords_value = configured.get("keywords")
    if isinstance(keywords_value, list):
        keywords = ", ".join(str(item) for item in keywords_value)
    elif keywords_value:
        keywords = str(keywords_value)
    else:
        terms = [item["word"] for item in guide["vocabulary"][:4]]
        keywords = ", ".join([str(guide["chinese_title"]), *map(str, terms)])
    return {
        "title": title,
        "author": author,
        "description": description[:150],
        "keywords": keywords,
    }


def list_items(items: list, *, ordered: bool = False) -> str:
    tag = "ol" if ordered else "ul"
    rows = "\n".join(f"    <li>{esc(item)}</li>" for item in items)
    return f"<{tag}>\n{rows}\n  </{tag}>"


def render_body(data: dict) -> str:
    guide = data["study_guide"]
    pages = clean_page_label(guide["source_pages"])
    paragraph_range = esc(guide["paragraph_range"])
    article_title = esc(data["title"])
    owner = esc(data["roles"]["word_master"]["name"])
    assignment_date = str(data.get("date") or datetime.date.today().isoformat())

    structure_rows = []
    for part in guide["structure_parts"]:
        page_range = clean_page_label(part["page_range"])
        structure_rows.append(
            "      <tr>"
            f"<td>第 {esc(page_range)} 页 · {esc(part['paragraph_range'])}</td>"
            f"<td>{esc(part['function'])}</td>"
            f"<td>{esc(part['content'])}</td>"
            "</tr>"
        )

    translation_rows = []
    for item in guide["translations"]:
        class_attr = ' class="page-break"' if bool(item.get("page_break_before")) else ""
        translation_rows.append(
            f"    <p{class_attr}><span class=\"para-label\">{esc(item['paragraph'])}</span> "
            f"{esc(item['text'])}</p>"
        )

    vocabulary_rows = []
    for item in guide["vocabulary"]:
        vocabulary_rows.append(
            "      <tr>"
            f"<td><strong>{esc(item['word'])}</strong><br>{esc(item['pronunciation'])} · {esc(item['part_of_speech'])}</td>"
            f"<td>{esc(item['meaning_zh'])}</td>"
            f"<td>{esc(item['context'])}</td>"
            "</tr>"
        )

    phrase_rows = "\n".join(
        f"    <li><span class=\"hl\">{esc(item['phrase'])}</span>：{esc(item['meaning_zh'])}</li>"
        for item in guide["additional_phrases"]
    )

    discussion_rows = "\n".join(
        f"    <p><strong>{index}.</strong> {esc(question)}</p>"
        for index, question in enumerate(guide["discussion_questions"], 1)
    )

    checklist_rows = []
    for item in guide["checklist"]:
        if isinstance(item, str):
            text = item
            checked = False
        else:
            text = item["text"]
            checked = bool(item.get("checked", False))
        mark = "☑" if checked else "□"
        checklist_rows.append(f"      <tr><td>{esc(text)}</td><td>{mark}</td></tr>")

    star = guide["star_word"]
    return f"""
<section class="cover">
  <div>
    <div class="cover-eyebrow">英语阅读圈 · {esc(guide['week_label'])}</div>
    <div class="cover-title">{esc(guide['chinese_title'])}</div>
    <div class="cover-sub">{article_title} 的中文精读与复习资料</div>
  </div>
  <div class="cover-meta">
    <strong>English Reading Circle</strong><br>
    Article study guide · {esc(assignment_date)}<br>
    基于教材第 {esc(pages)} 页文章整理
  </div>
</section>

<section class="toc">
  <h2>目录</h2>
  <div class="toc-item"><span class="toc-num">01</span><a class="toc-title" href="#ch-overview">文章主旨与结构</a></div>
  <div class="toc-item"><span class="toc-num">02</span><a class="toc-title" href="#ch-translation">{paragraph_range} 段落翻译</a></div>
  <div class="toc-item"><span class="toc-num">03</span><a class="toc-title" href="#ch-vocabulary">核心词汇与短语</a></div>
  <div class="toc-item"><span class="toc-num">04</span><a class="toc-title" href="#ch-review">复习清单与讨论</a></div>
</section>

<section class="chapter" id="ch-overview">
  <div class="chapter-num">01 · Overview</div>
  <h1>文章主旨与结构</h1>
  <p class="lead">{esc(guide['overview'])}</p>

  <h2>中心论点</h2>
  <div class="callout">{esc(guide['central_claim'])}</div>

  <h2>四段式逻辑</h2>
  <table class="kami-table striped">
    <thead><tr><th>页码与段落</th><th>作用</th><th>关键内容</th></tr></thead>
    <tbody>
{chr(10).join(structure_rows)}
    </tbody>
  </table>

  <h2>阅读提示</h2>
  {list_items(guide['reading_tips'])}
</section>

<section class="chapter" id="ch-translation">
  <div class="chapter-num">02 · Paragraph Translation</div>
  <h1>{paragraph_range} 段落翻译</h1>
  <p class="source-note">段落编号对应教材第 {esc(pages)} 页的原文编号。</p>
  <div class="translation">
{chr(10).join(translation_rows)}
  </div>
</section>

<section class="chapter" id="ch-vocabulary">
  <div class="chapter-num">03 · Vocabulary</div>
  <h1>核心词汇与短语</h1>
  <p class="lead">以下词汇来自 Word Master {owner} 的词汇表。先根据原句猜词义，再用释义核对。</p>
  <table class="kami-table compact striped">
    <thead><tr><th>词汇 / 发音 / 词性</th><th>中文义</th><th>原文语境</th></tr></thead>
    <tbody>
{chr(10).join(vocabulary_rows)}
    </tbody>
  </table>

  <div class="callout">
    <div class="chapter-num">Star Word · {esc(star['word'])}</div>
    <p>记忆画面：{esc(star['memory_picture_zh'])}</p>
    <p>例句：<span class="hl">{esc(star['example'])}</span></p>
    <p>思考：{esc(star['question_zh'])}</p>
  </div>

  <h2>补充短语</h2>
  <ul>
{phrase_rows}
  </ul>
</section>

<section class="chapter" id="ch-review">
  <div class="chapter-num">04 · Review</div>
  <h1>复习清单与讨论</h1>
  <h2>快速复述</h2>
  {list_items(guide['review_prompts'], ordered=True)}

  <h2>小组讨论问题</h2>
  <div class="callout">
{discussion_rows}
  </div>

  <h2>学习自检</h2>
  <table class="kami-table striped">
    <thead><tr><th>我能做到</th><th>完成</th></tr></thead>
    <tbody>
{chr(10).join(checklist_rows)}
    </tbody>
  </table>

  <div class="takeaway">
    <div class="takeaway-label">一句话带走</div>
    {esc(guide['takeaway'])}
  </div>
</section>
""".strip()


def render_html(data: dict, template_path: Path) -> str:
    if not template_path.is_file():
        fail(f"Study-guide template not found: {template_path}")
    template = template_path.read_text(encoding="utf-8")
    meta = metadata_for(data)
    replacements = {
        "{{DOC_TITLE}}": esc(meta["title"]),
        "{{AUTHOR}}": esc(meta["author"]),
        "{{DESCRIPTION}}": esc(meta["description"]),
        "{{KEYWORDS}}": esc(meta["keywords"]),
        "{{FOOTER_TITLE_CSS}}": css_string(data["study_guide"]["chinese_title"]),
        "{{BODY_CONTENT}}": render_body(data),
    }
    for marker, value in replacements.items():
        template = template.replace(marker, value)
    remaining = sorted(set(re.findall(r"\{\{[^{}]+\}\}", template)))
    if remaining:
        fail(f"Unresolved template markers: {', '.join(remaining)}")
    return template


def find_chrome() -> str:
    configured = os.environ.get("CHROME_BIN")
    if configured and Path(configured).is_file():
        return configured
    candidates = [
        Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
        Path("/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"),
        Path("/usr/bin/google-chrome"),
        Path("/usr/bin/chromium"),
        Path("/usr/bin/chromium-browser"),
    ]
    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)
    for name in ("google-chrome", "chromium", "chromium-browser", "microsoft-edge"):
        resolved = shutil.which(name)
        if resolved:
            return resolved
    fail("Chrome or Chromium is required for PDF export; use --html-only when unavailable")


def export_pdf(html_path: Path, pdf_path: Path, metadata: dict[str, str]) -> None:
    chrome = find_chrome()
    with tempfile.TemporaryDirectory(prefix="english_role_guide_") as temp_dir:
        temp = Path(temp_dir)
        raw_pdf = temp / "guide.pdf"
        profile = temp / "chrome-profile"
        command = [
            chrome,
            "--headless=new",
            "--disable-gpu",
            "--disable-background-networking",
            "--disable-component-update",
            "--no-pdf-header-footer",
            *(["--no-sandbox"] if hasattr(os, "geteuid") and os.geteuid() == 0 else []),
            "--run-all-compositor-stages-before-draw",
            f"--user-data-dir={profile}",
            f"--print-to-pdf={raw_pdf}",
            html_path.resolve().as_uri(),
        ]
        process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            start_new_session=(os.name != "nt"),
        )
        deadline = time.monotonic() + 45
        last_size = -1
        stable_reads = 0
        try:
            while time.monotonic() < deadline:
                if raw_pdf.is_file():
                    size = raw_pdf.stat().st_size
                    if size > 0 and size == last_size:
                        stable_reads += 1
                    else:
                        stable_reads = 0
                    last_size = size
                    if stable_reads >= 2:
                        break
                if process.poll() is not None and not raw_pdf.is_file():
                    break
                time.sleep(0.25)
        finally:
            if process.poll() is None:
                try:
                    if os.name == "nt":
                        process.terminate()
                    else:
                        os.killpg(process.pid, signal.SIGTERM)
                except ProcessLookupError:
                    pass
                try:
                    process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    try:
                        if os.name == "nt":
                            process.kill()
                        else:
                            os.killpg(process.pid, signal.SIGKILL)
                    except ProcessLookupError:
                        pass
                    process.wait(timeout=5)

        if not raw_pdf.is_file() or raw_pdf.stat().st_size == 0:
            stdout, stderr = process.communicate()
            detail = (stderr or stdout).strip()
            fail(f"Chrome PDF export failed: {detail[-1000:]}")

        reader = PdfReader(raw_pdf)
        writer = PdfWriter()
        writer.clone_document_from_reader(reader)
        writer.add_metadata(
            {
                "/Title": metadata["title"],
                "/Author": metadata["author"],
                "/Subject": metadata["description"],
                "/Keywords": metadata["keywords"],
            }
        )
        with pdf_path.open("wb") as stream:
            writer.write(stream)


def build(
    content_path: Path,
    output_dir: Path,
    template_path: Path = DEFAULT_TEMPLATE,
    *,
    html_only: bool = False,
) -> list[Path]:
    data = json.loads(content_path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        fail("Content JSON root must be an object")
    validate(data)
    output_dir.mkdir(parents=True, exist_ok=True)

    basename = output_basename(data)
    html_path = output_dir / f"{basename}.html"
    html_path.write_text(render_html(data, template_path), encoding="utf-8")
    outputs = [html_path]

    if not html_only:
        pdf_path = output_dir / f"{basename}.pdf"
        export_pdf(html_path, pdf_path, metadata_for(data))
        outputs.append(pdf_path)
    return outputs


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--content", type=Path, help="Combined UTF-8 assignment JSON")
    parser.add_argument("--output-dir", type=Path, help="Directory for HTML and PDF")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE, help="Override the HTML template")
    parser.add_argument("--html-only", action="store_true", help="Generate HTML without exporting PDF")
    parser.add_argument("--print-schema", action="store_true", help="Print the study-guide JSON schema")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.print_schema:
        print(json.dumps(SCHEMA, ensure_ascii=False, indent=2))
        return 0
    if args.content is None or args.output_dir is None:
        print("error: --content and --output-dir are required unless --print-schema is used", file=sys.stderr)
        return 2
    try:
        outputs = build(
            args.content.resolve(),
            args.output_dir.resolve(),
            args.template.resolve(),
            html_only=args.html_only,
        )
    except (OSError, ValueError, json.JSONDecodeError, subprocess.SubprocessError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    for output in outputs:
        print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
