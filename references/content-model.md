# Assignment Content Model

Use one UTF-8 JSON file for one article. The Python scripts are stable; edit the JSON, not the scripts or generated files. Start from the complete runnable [assignment.example.json](../assets/assignment.example.json), copy it, and replace the example article content.

Print the complete machine-readable schema with:

```bash
python3 scripts/build_assignment.py --print-schema
```

## Shared fields

- `title`: full English article title.
- `date`: assignment date in `YYYY-MM-DD` format.
- `roles`: content for the four DOCX worksheets.
- `study_guide`: content for the Kami Chinese HTML and PDF.

Each role requires `name`, which controls the output filename. When a Chinese name may not render reliably in the DOCX QA renderer, also provide `display_name` with its English or romanized form; this controls the visible worksheet header without changing the filename.

Every source location must use the printed textbook page, such as `Page 31, P7` or `Pages 30-31, P1-P4`. Do not use image filenames.

## Checkbox fields

Checkboxes are selected by stable keys. Remove a key to leave that option unchecked.

Discussion Leader `answerers` (the "Who answered?" column, filled with full names): when `meeting_completed` is true, blank entries are assigned automatically to the other group members in a stable pseudo-random order; supply names only to override.

Discussion Leader `participation`:

- `everyone_spoke`
- `stayed_on_topic`
- `connected_to_experience`
- `ended_with_big_question`

When `participation` is omitted, all four are checked if `meeting_completed` is true and none are checked otherwise.

Summarizer `purposes`:

- `describe`
- `explain_why`
- `explain_steps`
- `compare_contrast`

Summarizer `structures`:

- `time_order`
- `space_order`
- `main_idea_details`
- `cause_effect`
- `compare_contrast`

Word Master `categories`:

- `new`
- `striking`
- `repeated`
- `specialist`
- `idioms`

Structure Analyzer `structures`:

- `story_plot`
- `sequence`
- `cause_effect`
- `compare_contrast`
- `problem_solution`
- `description_list`

Chinese study-guide checklist items may be plain strings, which render unchecked, or objects:

```json
{"text": "我能复述文章的主要步骤。", "checked": true}
```

## Study guide fields

The guide uses four fixed chapters: overview, paragraph translation, vocabulary, and review. Supply:

- printed `source_pages` and `paragraph_range`;
- exactly four `structure_parts` with page and paragraph ranges;
- every translated paragraph with its printed `page`;
- four to eight vocabulary entries;
- one Star Word and at least one additional phrase;
- at least three review prompts, exactly four discussion questions, and at least three checklist items.

Set `page_break_before: true` on a translation only when a visual QA pass shows that the next paragraph should begin on a new page. `metadata` and `output_basename` are optional; the builder derives sensible defaults.

## Commands

Build and verify everything:

```bash
python3 scripts/build_assignment.py \
  --content /absolute/path/to/assignment.json \
  --output-dir /absolute/path/to/output
```

Recheck existing files without rebuilding:

```bash
python3 scripts/build_assignment.py \
  --content /absolute/path/to/assignment.json \
  --output-dir /absolute/path/to/output \
  --verify-only
```

Use `--html-only` only when Chrome or Chromium is unavailable. The normal workflow must deliver both HTML and PDF.
