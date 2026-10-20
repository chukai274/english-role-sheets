---
name: english-role-sheets
description: Read an English article from images, PDF, DOCX, or text; collect the four reading-circle role owners; then use a reusable JSON-driven pipeline to produce four completed Role Work Sheets DOCX files plus a Kami-formatted Chinese study guide. Use for four-person English reading assignments with Discussion Leader, Summarizer, Word Master, and Structure Analyzer roles.
---

# English Role Sheets

Turn one English reading into five verified deliverables:

1. one DOCX for the Discussion Leader;
2. one DOCX for the Summarizer;
3. one DOCX for the Word Master;
4. one DOCX for the Structure Analyzer;
5. one Chinese study guide in Kami HTML and PDF.

All assignment-specific text, page references, owners, and checkbox choices belong in one UTF-8 JSON file. Do not edit the Python scripts, DOCX template, HTML template, or generated files for a new week. Read [references/content-model.md](references/content-model.md) when preparing or revising the JSON.

## Intake gate

Collect information in this order. Do not begin authoring before both gates are complete.

1. If the article source is not already explicit, ask: `文章材料在哪里？请给我文件或文件夹路径。`
2. Resolve and inspect the source. A folder may contain ordered page images, a PDF, DOCX, or text file. Sort page images by their numeric filename parts. Ignore worksheets, prior outputs, temporary Office files, and QA folders.
3. If the source is unreadable or page order is genuinely ambiguous, ask one focused question. Otherwise identify the title and confirm that the full article is available.
4. Ask for all four owners in one message:

   ```text
   请告诉我四位负责人：
   Discussion Leader：
   Summarizer：
   Word Master：
   Structure Analyzer：
   ```

Skip any question already answered by the user. Use the current local date unless the user gives an assignment date.

## Read and ground the article

Read every page before drafting. Use the printed page number visible on the article whenever one exists, with stable anchors such as `Page 31, Paragraph 11`. Do not use image filenames (`Image 1`, `图片 1`) as citations when printed page numbers are available. Preserve names, numbers, quotations, and causal claims exactly. Never fill an unclear passage from guesswork; inspect it again or ask the user.

Use the bundled [Role work sheets.docx](assets/Role%20work%20sheets.docx) as the document template. Do not redesign it.

## Draft the four roles

Write clear, natural English appropriate for the source text.

### Discussion Leader

- Write exactly four open-ended questions that move from comprehension to analysis and application.
- Select one source passage and explain why it matters.
- Supply one best idea for each question. If the user requests a fully completed worksheet or confirms the discussion is complete, set `meeting_completed` to true: every `Who answered?` field gets one group member's full name (never initials; leave `answerers` entries blank to have the builder assign the other group members automatically) and all supported participation items are checked. Otherwise leave the names blank and mark the items for completion after the meeting.
- Do not claim a meeting happened unless the user requests the completed post-discussion form or confirms it happened.
- Use past tense for a completed form (`My strongest question was...`) and prospective language otherwise (`My strongest question will be...`).

### Summarizer

- State the topic in one sentence and mark every applicable author purpose.
- Divide the article into four logical parts. Give each part a source range, main idea, and one grounded fact.
- Identify all applicable organizational structures.
- Write a coherent full summary, followed by a second summary of no more than 40 English words.
- Choose two indispensable key terms and ensure both appear in the short summary.

### Word Master

- Select four to six useful words or phrases from the article; prefer recurring, topic-specific, or inference-worthy vocabulary.
- Copy each source sentence exactly, record its source anchor, give a plausible contextual guess, and then a concise dictionary-style meaning.
- Choose one Star Word. Add a memory-picture description, an original example sentence, and a discussion question.

### Structure Analyzer

- Mark every structure genuinely present, but identify the dominant structure in the proof sentence.
- Record signal words that occur in or accurately describe the article's organization.
- Build a four-part map with source ranges and concise explanations.
- Explain how the organization supports the meaning; do not merely repeat the structure label.
- If the group has not met, phrase the “surprised the group” field as a likely point for discussion rather than a completed event.

## Prepare the content JSON

Create one UTF-8 JSON file matching the combined schema printed by:

```bash
python3 scripts/build_assignment.py --print-schema
```

The checkbox arrays use stable keys documented in [references/content-model.md](references/content-model.md). Select only choices supported by the article or completed meeting. Keep every source reference tied to the printed page number.

## Build all deliverables

Run the single entry point:

```bash
python3 scripts/build_assignment.py \
  --content /absolute/path/to/assignment.json \
  --output-dir /absolute/path/to/output
```

This command validates the content, fills the fixed DOCX template, applies every checkbox state, builds the Summarizer Details Chart as a real table, creates the Kami HTML, exports the PDF through local Chrome or Chromium (set `CHROME_BIN` when the browser is in a non-standard location; `--no-sandbox` is added automatically when running as root), writes PDF metadata, and runs structural checks across all six files, including a check that no template blank line is left unfilled. Use a Python 3 interpreter that already has the packages in `requirements.txt` installed.

Use `build_role_docs.py`, `build_study_guide.py`, or `verify_outputs.py` directly only when diagnosing one stage. Do not hand-edit generated deliverables to fix a reusable behavior; correct the JSON for assignment content or the relevant script for pipeline behavior, then rebuild.

## Verify the four DOCX files

The unified builder runs structural checks first. Then load the installed `documents` skill and follow its render-and-inspect gate. Render every DOCX to page PNGs with its `render_docx.py`, inspect all pages at full size, and revise JSON content or reusable layout code if text clips, tables overflow, glyphs disappear, or pagination looks broken. QA renders are internal unless requested.

Also verify:

- each file contains only its assigned role;
- the owner, group, title, and date are correct;
- all article references resolve to printed page numbers and paragraph numbers where available, with no image-filename citations;
- the Summarizer short version is at most 40 words;
- no meeting activity has been fabricated.

## Create the Chinese study guide with Kami

Load the installed `kami` skill and follow its update, content, and visual QA requirements. The reusable `assets/study-guide-template.html` is the fixed Kami template; `build_study_guide.py` fills it from `study_guide` in the same JSON and exports HTML plus PDF.

The default audience is students using the material for preview, group discussion, and revision. Scale the document to the article instead of padding it; a normal target is 5-8 A4 pages.

Include, when supported by the article:

1. cover and contents;
2. concise Chinese overview and central claim;
3. article structure or step-by-step logic;
4. paragraph-numbered Chinese translation;
5. the Word Master vocabulary with pronunciation, part of speech, Chinese meaning, and source context;
6. additional important phrases;
7. short review checklist or comprehension prompts.

Use only facts from the source. Keep English-term density controlled, preserve printed page and paragraph numbers, and do not place the Chinese guide under any one role owner's name. The builder checks placeholders and metadata; render and inspect every PDF page before handoff.

## Handoff

Return five deliverable groups unless the user asks for more: four role DOCX files and one Chinese guide supplied as both HTML and PDF, for six user-facing files in total. List each file with its owner or purpose. Do not include intermediate JSON, rendered QA PNGs, or temporary render files.
