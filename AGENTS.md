# AGENTS.md

## Purpose

This repository maintains the `case-writer` Codex skill for writing finance-focused business school teaching cases and teaching notes from background materials.

## Dependency Check

Use the repository Python wrapper so the intended Conda environment is used even when the shell's `python` command points to the Windows Store launcher. Before extracting text, validating skill metadata, or producing intermediate documents, run:

```powershell
.\scripts\run_python.ps1 scripts/check_dependencies.py
```

By default, the wrapper uses `D:\anaconda3\envs\python313\python.exe`. To use a different interpreter temporarily, set `CASE_WRITER_PYTHON` to the full path of the desired `python.exe`.

The required Python packages are listed in `requirements.txt`:

- `PyMuPDF` for PDF text extraction through the `fitz` module.
- `PyYAML` for YAML/frontmatter checks and skill validation workflows.

The script prints the Python executable and version so environment mix-ups are easy to spot.

## Intermediate Text Products

Use the repository scripts for generated text artifacts instead of ad hoc extraction commands:

```powershell
.\scripts\run_python.ps1 scripts/extract_text.py supporting_documents/background --output outputs/text/background
```

The script supports `.pdf`, `.docx`, `.txt`, `.md`, and `.html`/`.htm` inputs. Case-writing intermediate products should stay under `outputs/`; do not use `artifacts/` for case-writing runs. Delete non-deliverable intermediate products, such as `outputs/text/`, at the end of the run unless the user explicitly asks to keep them.

## Plan-First Case Workflow

Before drafting a teaching case, ask the user what they want the case to emphasize: protagonist or decision maker, company/event scope, core problem, course/audience, teaching emphasis, and whether the case should be prospective or retrospective.

After source inventory and fact-base preparation, present a concise plan and wait for explicit user approval before writing any student-facing case draft, teaching note, substantial narrative section, or final PDF. Before approval, only dependency checks, source extraction, source inventory, short fact-base notes, clarification questions, and proposed plans are allowed.

For case writing, use `supporting_documents/examples/安德科铭/安德科铭正文案例.pdf` and `supporting_documents/examples/安德科铭/安德科铭案例使用说明.pdf` as the default format templates. The student-facing case body should be 8,000-10,000 Chinese characters unless the user explicitly changes the length. The teaching note (案例使用说明) should be 10,000-15,000 Chinese characters unless the user explicitly changes the length.

## Final Outputs

Current repository rule: final deliverables belong in `outputs/` and include both DOCX drafts and PDFs. Keep `outputs/案例正文.docx`, `outputs/案例正文.pdf`, `outputs/案例使用说明.docx`, and `outputs/案例使用说明.pdf`. Do not use `--delete-source` for final DOCX drafts. Any older examples in this file that mention deleting draft sources are superseded by this rule.

Final case deliverables belong in `outputs/` and should include both DOCX drafts and PDFs:

- `outputs/案例正文.pdf`
- `outputs/案例正文.docx`
- `outputs/案例使用说明.pdf`
- `outputs/案例使用说明.docx`

Both final PDFs should be written in Chinese unless another language is explicitly requested. Facts drawn from `supporting_documents/background/` should be marked with footer-style endnotes or a clearly labeled source endnote block when page footers are not technically available.

Draft in DOCX first, then convert to PDF with `scripts/write_pdf.py`. Keep both the DOCX drafts and PDFs:

```powershell
.\scripts\run_python.ps1 scripts/write_pdf.py outputs/案例正文.docx outputs/案例正文.pdf
.\scripts\run_python.ps1 scripts/write_pdf.py outputs/案例使用说明.docx outputs/案例使用说明.pdf
```

Keep the final DOCX drafts in `outputs/`. Delete only non-deliverable intermediate products after final PDFs are generated.

## Source Material Policy

- `supporting_documents/background/` contains public news, public company pages, and other source material for the case being written.
- `supporting_documents/examples/` contains representative teaching-case examples that informed the `case-writer` method.
- `supporting_documents/methodology/` contains case-writing methodology materials that informed the `case-writer` method.
- The `supporting_documents/` contents may be committed to GitHub for this repository.

## Editing Guidance

- Keep the skill itself in `.codex_skills/case-writer/`.
- Put reusable workflow knowledge in `.codex_skills/case-writer/references/` rather than making `SKILL.md` overly long.
- Keep skill references focused on reusable writing principles rather than duplicating source documents.
