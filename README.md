# english-role-sheets

四人英语阅读圈（Discussion Leader、Summarizer、Word Master、Structure Analyzer）作业生成 skill。

给它一篇英文课文和四位负责人的名字，它用一份 JSON 生成：

- 4 份角色作业 DOCX（每个角色一份，沿用 `assets/Role work sheets.docx` 模板）
- 1 份中文学习资料（HTML + PDF）

An AI-agent skill that turns one English reading into four role worksheets (DOCX) and a Chinese study guide (HTML + PDF), driven by a single JSON file.

## 使用

1. 安装依赖：`pip install -r requirements.txt`，并安装 Chrome 或 Chromium（导出 PDF 用；浏览器在非常规位置时设置环境变量 `CHROME_BIN`）。
2. 复制 `assets/assignment.example.json`，换成自己的课文内容（字段说明见 `references/content-model.md`，完整结构用 `--print-schema` 打印）。
3. 生成并校验全部文件：

```bash
python3 scripts/build_assignment.py \
  --content /path/to/assignment.json \
  --output-dir /path/to/output
```

`SKILL.md` 是给 AI 助手读的完整工作流（取材、写内容、生成、检查、交付）。

## 目录

- `SKILL.md`：skill 说明与工作流
- `scripts/`：`build_assignment.py`（入口）、`build_role_docs.py`、`build_study_guide.py`、`verify_outputs.py`
- `assets/`：DOCX 模板、中文学习资料 HTML 模板、示例 JSON
- `references/content-model.md`：JSON 字段与勾选框 key 说明

## 行为说明

- 勾选框只按 JSON 里给出的 key 勾选，没给的保持不勾。
- Discussion Leader：`meeting_completed` 为 false 时，"Who answered?"留空、参与情况不勾；为 true 时，空着的 `answerers` 会自动用其他组员的全名补上，参与情况全勾。
- Summarizer 的 Details Chart 生成为真表格。
- 校验会检查勾选框数量、空栏目、模板残留的下划线、页码引用等。
- 以 root 身份运行时导出 PDF 会自动给 Chrome 加 `--no-sandbox`。

## 许可

尚未指定许可证。`assets/Role work sheets.docx` 是课程作业表模板，使用或再分发前请确认其来源允许。
