# Case Study Writer

This repository contains the `case-writer` Codex skill and supporting scripts for producing finance-focused business school teaching cases from background materials.

The skill is designed for undergraduate and graduate business education. Given a background folder about a company, entrepreneur, financing event, or strategic finance decision, it guides Codex to produce two independent PDF files:

1. A student-facing teaching case narrative.
2. An instructor-facing case teaching note with discussion-question reference answers.

## Repository Structure

```text
.codex_skills/
  case-writer/
    SKILL.md
    agents/openai.yaml
    references/
scripts/
  check_dependencies.py
  extract_text.py
  write_pdf.py
outputs/
  .gitkeep
supporting_documents/
  background/
  examples/
  methodology/
```

## Source Material Policy

`supporting_documents/` stores the materials used to build and run the case-writing workflow:

- `supporting_documents/background/` contains public background materials for the case being written.
- `supporting_documents/examples/` contains representative teaching-case examples that informed the `case-writer` method.
- `supporting_documents/methodology/` contains case-writing methodology materials that informed the `case-writer` method.

The `supporting_documents/` contents may be committed to GitHub for this repository.

## Setup

This repository is configured to use the Conda environment at:

```text
D:\anaconda3\envs\python313\python.exe
```

VS Code points to that interpreter through `.vscode/settings.json`, and PowerShell commands should use the repository wrapper `scripts/run_python.ps1` so they do not accidentally hit the Windows Store `python.exe` launcher. To override the interpreter for one shell session, set `CASE_WRITER_PYTHON` to a different full `python.exe` path.

Install dependencies with:

```powershell
D:\anaconda3\envs\python313\Scripts\pip.exe install -r requirements.txt
```

Check that required packages are available:

```powershell
.\scripts\run_python.ps1 scripts/check_dependencies.py
```

The script prints the Python executable and version so you can confirm it is using the intended environment.

The project expects:

- `PyMuPDF`, imported as `fitz`, for PDF extraction and PDF output.
- `PyYAML`, imported as `yaml`, for YAML/frontmatter checks.

## Generate Intermediate Text

Use the internal extraction script to convert background materials into text artifacts:

```powershell
.\scripts\run_python.ps1 scripts/extract_text.py supporting_documents/background --output outputs/text/background
```

Generated text for case-writing runs should be written under `outputs/`, not `artifacts/`. Delete non-deliverable intermediate products, such as `outputs/text/`, at the end of the run unless you explicitly choose to keep them.

## Using The Skill

The skill source lives at:

```text
.codex_skills/case-writer/SKILL.md
```

Typical invocation:

```text
Use $case-writer to write a finance-focused business teaching case from supporting_documents/background.
```

The skill follows a plan-first workflow. Before drafting, it must ask the user to confirm the desired protagonist or decision maker, company/event scope, core problem, course/audience, teaching emphasis, and whether the case should be prospective or retrospective. After source preparation, it must present a concise writing plan and wait for explicit user approval before creating any student case draft, teaching note, substantial narrative section, or final PDF.

The default format templates are `supporting_documents/examples/安德科铭/安德科铭正文案例.pdf` and `supporting_documents/examples/安德科铭/安德科铭案例使用说明.pdf`. The student-facing case body should be 8,000-10,000 Chinese characters unless the user explicitly changes the length.

The expected final outputs are:

- `outputs/案例正文.md`
- `outputs/案例正文.pdf`
- `outputs/案例使用说明.md`
- `outputs/案例使用说明.pdf`

Both final PDFs should be written in Chinese unless another language is explicitly requested. Facts drawn from `supporting_documents/background/` should be marked with footer-style endnotes or a clearly labeled source endnote block when page footers are not technically available.

Draft in Markdown first, then render it to PDF. Keep both the Markdown and PDF files:

```powershell
.\scripts\run_python.ps1 scripts/write_pdf.py outputs/案例正文.md outputs/案例正文.pdf
.\scripts\run_python.ps1 scripts/write_pdf.py outputs/案例使用说明.md outputs/案例使用说明.pdf
```

Generated files inside `outputs/` are ignored by Git except for `.gitkeep`.

## Publishing To GitHub

Before pushing:

```powershell
git status --ignored
```

Confirm that the intended `supporting_documents/` files are staged, because this repository is configured to allow uploading them.
