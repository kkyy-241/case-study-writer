# AGENTS.md

## Purpose

This repository maintains the `case-writer` Codex skill for writing finance-focused business school teaching cases and teaching notes from background materials.

## Dependency Check

Use the repository Python wrapper so the intended Conda environment is used even when the shell's `python` command points to the Windows Store launcher. Before extracting text, validating skill metadata, or producing intermediate documents, run:

```powershell
.\scripts\run_python.ps1 scripts/check_dependencies.py
```

By default, the wrapper uses this workstation's Conda interpreter at `D:\anaconda3\envs\python313\python.exe`. Other users should set `CASE_WRITER_PYTHON` to the full path of their desired `python.exe` before running repository scripts.

The required Python packages are listed in `requirements.txt`:

- `PyMuPDF` for PDF text extraction through the `fitz` module.
- `PyYAML` for YAML/frontmatter checks and skill validation workflows.
- `python-docx` for DOCX draft generation through the `docx` module.
- `pywin32` on Windows for Microsoft Word COM PDF export through the `win32com` module.

Microsoft Word with `pywin32` is recommended on Windows for the most faithful DOCX-to-PDF conversion. LibreOffice is the cross-platform fallback. Without either, `scripts/write_pdf.py` falls back to a lightweight PyMuPDF text renderer.

The script prints the Python executable and version so environment mix-ups are easy to spot.

## Intermediate Text Products

Use the repository scripts for generated text artifacts instead of ad hoc extraction commands:

```powershell
.\scripts\run_python.ps1 scripts/extract_text.py supporting_documents/background --output outputs/text/background
```

The script supports `.pdf`, `.docx`, `.txt`, `.md`, and `.html`/`.htm` inputs. Extracted source text and other temporary products should stay under `outputs/text/`; do not use `artifacts/` for case-writing runs. Delete temporary products such as `outputs/text/` at the end of the run unless the user explicitly asks to keep them.

Retain build drafts under `outputs/drafts/`. The UTF-8 draft files used to build DOCX deliverables are not temporary products; keep them so a user can revise text and regenerate DOCX/PDF without rerunning the whole case workflow.

## Image Assets

Use repository scripts for generated image artifacts instead of ad hoc extraction or download commands. Before drafting a case that would benefit from visuals, collect candidate images from the background folder:

```powershell
.\scripts\run_python.ps1 scripts/collect_images.py supporting_documents/background --output outputs/drafts/images
```

The script extracts images from PDFs and DOCX files, copies local image files, discovers local images referenced by HTML files, and writes `manifest.json` and `manifest.md` with traceable source notes.

If `supporting_documents/background/` does not contain suitable images, use web/image search to identify appropriate public images, then download only the selected direct image URLs with source notes:

```powershell
.\scripts\run_python.ps1 scripts/collect_images.py `
  --url "https://example.com/path/image.jpg" `
  --source-note "Image title, publisher/site, publication date or access date, source page URL" `
  --output outputs/drafts/images
```

Draft sources may reference collected images with Markdown image syntax such as `![图1 股权结构示意图](images/ownership-chart.png)`. `scripts/write_docx.py` embeds these images in DOCX. Final PDFs should be generated from DOCX through Word COM. Direct text-to-PDF rendering is only an emergency fallback after telling the user that its layout will not fully match the DOCX.

Only use images that support case analysis or classroom exhibits. Preserve image source traceability in the case source notes or in the generated image manifest. The workflow must never create its own charts, diagrams, or synthetic images; visuals may only be extracted from `supporting_documents/background/` or collected through web/image search.

## Archive Before New Case Runs

Every time the user asks to write a new case, confirm the `<CompanyName>` suffix and ask whether to archive the current `supporting_documents/background/` and `outputs/` files into `outputs/archive/<CompanyName>/` before drafting. The archive folder name must exactly match the company suffix used in generated deliverable filenames. For example, if the confirmed suffix is `ExampleCo`, files such as `案例正文_ExampleCo.docx` correspond to `outputs/archive/ExampleCo/`.

Use the repository script after the user approves archiving:

```powershell
.\scripts\run_python.ps1 scripts/archive_case.py ExampleCo --mode copy
```

Use `--mode move` only when the user explicitly wants to clear the active background/output folders for the next run.

## Plan-First Case Workflow

Before drafting a teaching case, ask the user what they want the case to emphasize: protagonist or decision maker, company/event scope, core problem, course/audience, teaching emphasis, and whether the case should be prospective or retrospective.

During the plan stage, explicitly ask for and confirm the exact `<CompanyName>` suffix used in retained drafts, final deliverables, and archive folders. If the user has not provided it, propose one based on the case subject and ask for confirmation. Do not silently infer the suffix from file names, background materials, or prior case examples. Do not mention any prior case company name when asking for the suffix.

Generated DOCX/PDF page headers must use `商学院教学案例库` left-aligned and the confirmed `<CompanyName>` suffix right-aligned. The right header text must match the filename suffix and archive folder name.

After source inventory and fact-base preparation, present a concise plan and wait for explicit user approval before writing any student-facing case draft, teaching note, substantial narrative section, or final PDF. The plan must include the confirmed `<CompanyName>` suffix and archive plan. Before approval, only dependency checks, source extraction, source inventory, short fact-base notes, clarification questions, and proposed plans are allowed.

For case writing, use `.codex_skills/case-writer/references/format-templates/安德科铭正文案例.pdf` and `.codex_skills/case-writer/references/format-templates/安德科铭案例使用说明.pdf` as the default format templates. The student-facing case body should be 8,000-10,000 Chinese characters unless the user explicitly changes the length. The teaching note (案例使用说明) should be 10,000-15,000 Chinese characters unless the user explicitly changes the length.

## Final Outputs

Current repository rule: final deliverables belong in `outputs/` and include both DOCX drafts and PDFs. File names must end with the related company name suffix, matching the archive folder name, such as `outputs/案例正文_<CompanyName>.docx`, `outputs/案例正文_<CompanyName>.pdf`, `outputs/案例使用说明_<CompanyName>.docx`, and `outputs/案例使用说明_<CompanyName>.pdf`. Keep final DOCX drafts after PDF conversion. Any older examples in this file that mention deleting draft sources are superseded by this rule.

Final case deliverables belong in `outputs/` and should include both DOCX drafts and PDFs:

- `outputs/案例正文_<CompanyName>.pdf`
- `outputs/案例正文_<CompanyName>.docx`
- `outputs/案例使用说明_<CompanyName>.pdf`
- `outputs/案例使用说明_<CompanyName>.docx`

Both final PDFs should be written in Chinese unless another language is explicitly requested. Facts drawn from `supporting_documents/background/` should be marked with footer-style endnotes or a clearly labeled source endnote block when page footers are not technically available.

After approval, create retained UTF-8 build drafts under `outputs/drafts/`, convert them to DOCX with `scripts/write_docx.py`, then convert the DOCX files to PDF with `scripts/write_pdf.py --backend word`. If Word COM fails, stop and tell the user instead of silently falling back to PyMuPDF. Keep the retained text drafts, DOCX drafts, and PDFs:

```powershell
.\scripts\run_python.ps1 scripts/write_docx.py outputs/drafts/案例正文_ExampleCo.txt outputs/案例正文_ExampleCo.docx
.\scripts\run_python.ps1 scripts/write_docx.py outputs/drafts/案例使用说明_ExampleCo.txt outputs/案例使用说明_ExampleCo.docx
.\scripts\run_python.ps1 scripts/write_pdf.py outputs/案例正文_ExampleCo.docx outputs/案例正文_ExampleCo.pdf --backend word
.\scripts\run_python.ps1 scripts/write_pdf.py outputs/案例使用说明_ExampleCo.docx outputs/案例使用说明_ExampleCo.pdf --backend word
```

Keep the final DOCX drafts in `outputs/` and the retained text drafts in `outputs/drafts/`. Delete only non-deliverable temporary products, including `outputs/text/`, after final PDFs are generated.

## Source Material Policy

- `supporting_documents/background/` contains public news, public company pages, and other source material for the case being written.
- `supporting_documents/examples/` contains representative teaching-case examples used, through Codex's built-in `skill-creator` workflow, to derive the `case-writer` skill's default case structure, document style, template expectations, and quality checks.
- `supporting_documents/methodology/` contains case-writing methodology materials used, through Codex's built-in `skill-creator` workflow, to derive the `case-writer` skill's writing principles, plan-first workflow, teaching-note requirements, and evaluation standards.
- `.codex_skills/case-writer/references/format-templates/` contains the default runtime format templates inspected by `case-writer`; keep template files there rather than mixing them into user-collected examples.
- Users of this repository may replace or extend `supporting_documents/examples/` with their own representative cases and `supporting_documents/methodology/` with their own case-writing requirements, then use `skill-creator` to update `.codex_skills/case-writer/` so the skill reflects those materials.
- `supporting_documents/background/` contents may be committed when they are public source materials intended to accompany the case-writing run. `supporting_documents/examples/` and `supporting_documents/methodology/` are ignored by default because they are user-collected materials for building or updating the skill.

## Editing Guidance

- Keep the skill itself in `.codex_skills/case-writer/`.
- Put reusable workflow knowledge in `.codex_skills/case-writer/references/` rather than making `SKILL.md` overly long.
- Keep skill references focused on reusable writing principles rather than duplicating source documents.
