# Case Study Writer

This repository contains the `case-writer` Codex skill and supporting scripts for producing finance-focused business school teaching cases from background materials.

The skill is designed for undergraduate and graduate business education. Given a background folder about a company, entrepreneur, financing event, or strategic finance decision, it guides Codex to produce two independent DOCX drafts and two independent PDF files:

1. A student-facing teaching case narrative.
2. An instructor-facing case teaching note with discussion-question reference answers.

## Repository Structure

```text
.codex_skills/
  case-writer/
    SKILL.md
    agents/openai.yaml
    references/
      format-templates/
scripts/
  check_dependencies.py
  collect_images.py
  archive_case.py
  extract_text.py
  write_docx.py
  write_pdf.py
outputs/
  .gitkeep
examples/
  generated/
supporting_documents/
  background/
  examples/
  methodology/
```

## Example Outputs

Sample generated PDFs are committed under `examples/generated/` to show the repository's output style.

Runtime outputs are written to `outputs/` and are ignored by Git by default, including generated DOCX and PDF files.

## Source Material Policy

`supporting_documents/` stores the materials used to build and run the case-writing workflow:

- `supporting_documents/background/` contains public background materials for the case being written.
- `supporting_documents/examples/` contains representative teaching-case examples used as source material for building or updating the `case-writer` skill with `skill-creator`; these are not runtime templates.
- `supporting_documents/methodology/` contains case-writing methodology materials used as source material for building or updating the `case-writer` skill with `skill-creator`.
- `.codex_skills/case-writer/references/format-templates/` contains the default runtime format templates that the skill inspects when drafting case deliverables.

## How The Skill Was Built And Updated

The `case-writer` skill was created with Codex's built-in `skill-creator` workflow. Its writing workflow, structure references, and quality checks were derived from the teaching-case examples collected in `supporting_documents/examples/` and the case-writing methodology articles collected in `supporting_documents/methodology/`.

Users of this repository can build or update `case-writer` with their own materials. Replace or add to `supporting_documents/examples/` with representative cases, replace or add to `supporting_documents/methodology/` with preferred case-writing requirements, then ask Codex to use `skill-creator` to revise `.codex_skills/case-writer/` so the skill reflects those examples and standards.

## Setup

This repository's current workstation default uses the Conda environment at:

```text
D:\anaconda3\envs\python313\python.exe
```

VS Code points to that interpreter through `.vscode/settings.json`, and PowerShell commands should use the repository wrapper `scripts/run_python.ps1` so they do not accidentally hit the Windows Store `python.exe` launcher. Other users should set `CASE_WRITER_PYTHON` to their own full `python.exe` path before running repository scripts.

Example override for one PowerShell session:

```powershell
$env:CASE_WRITER_PYTHON="C:\Path\To\python.exe"
.\scripts\run_python.ps1 scripts/check_dependencies.py
```

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
- `python-docx`, imported as `docx`, for DOCX draft generation.
- Microsoft Word with `pywin32` is recommended on Windows for the most faithful DOCX-to-PDF conversion. LibreOffice is the cross-platform fallback. Without either, `scripts/write_pdf.py` falls back to a lightweight PyMuPDF text renderer.

## Generate Intermediate Text

Use the internal extraction script to convert background materials into text artifacts:

```powershell
.\scripts\run_python.ps1 scripts/extract_text.py supporting_documents/background --output outputs/text/background
```

Generated extracted text for case-writing runs should be written under `outputs/text/`. Delete temporary extraction products, such as `outputs/text/`, at the end of the run unless you explicitly choose to keep them.

Build drafts used to create DOCX files should be written under `outputs/drafts/` and retained. This lets you edit the draft text and regenerate DOCX/PDF deliverables without rerunning the full case-writing workflow.

## Collect Case Images

Case drafts may include image syntax in the UTF-8 draft source:

```markdown
![图1 股权结构示意图](../images/ownership-chart.png)
```

`scripts/write_docx.py` embeds those images in the DOCX. Final PDFs should be generated from DOCX through Microsoft Word COM for the closest match to DOCX layout. Direct text-to-PDF rendering is only an emergency fallback and will not fully match the DOCX layout.

Before drafting, collect usable image assets from `supporting_documents/background/`:

```powershell
.\scripts\run_python.ps1 scripts/collect_images.py supporting_documents/background --output outputs/drafts/images
```

The script extracts images from PDF and DOCX files, copies existing image files, discovers local images referenced by HTML files, and writes `manifest.json` plus `manifest.md` with source traceability.

If the background materials do not contain suitable images for exhibits such as product photos, headquarters photos, founder portraits, financing diagrams, market screenshots, or ownership visuals, use web/image search to find appropriate public images. Then download the selected direct image URLs with source notes:

```powershell
.\scripts\run_python.ps1 scripts/collect_images.py `
  --url "https://example.com/path/image.jpg" `
  --source-note "Image title, publisher/site, publication date or access date, source page URL" `
  --output outputs/drafts/images
```

Only use images that are relevant to the case analysis or classroom exhibits, and preserve the source note in the case's source notes or image manifest. The workflow must not create its own diagrams, charts, or synthetic images; visuals must come from `supporting_documents/background/` or from web/image search.

## Archive Completed Runs

Before starting a new case, confirm the `<CompanyName>` suffix and ask whether to archive the current `supporting_documents/background/` and `outputs/` files into `outputs/archive/<CompanyName>/`. The archive folder name must match the company suffix used in the generated deliverable filenames.

Example:

```powershell
.\scripts\run_python.ps1 scripts/archive_case.py ExampleCo --mode copy
```

This creates `outputs/archive/ExampleCo/`. Use `--mode move` only when you explicitly want to clear the active background/output folders for the next case.

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

During the plan stage, the workflow must also confirm the exact `<CompanyName>` suffix used in retained drafts, final deliverables, and archive folders. If the user has not provided it, Codex should propose one based on the case subject and ask for confirmation rather than silently inferring it from filenames, background materials, or prior case examples. The prompt should not mention any prior case company name.

Generated DOCX/PDF page headers use `商学院教学案例库` left-aligned and the confirmed `<CompanyName>` suffix right-aligned. The right header text must match the filename suffix and archive folder name.

The default format templates are `.codex_skills/case-writer/references/format-templates/安德科铭正文案例.pdf` and `.codex_skills/case-writer/references/format-templates/安德科铭案例使用说明.pdf`. The student-facing case body should be 8,000-10,000 Chinese characters unless the user explicitly changes the length. The teaching note ranges between 10,000-15,000 Chinese characters unless the user explicitly changes the length.

The expected final outputs include the related company name as a suffix:

- `outputs/案例正文_<CompanyName>.docx`
- `outputs/案例正文_<CompanyName>.pdf`
- `outputs/案例使用说明_<CompanyName>.docx`
- `outputs/案例使用说明_<CompanyName>.pdf`

Both final PDFs should be written in Chinese unless another language is explicitly requested. Facts drawn from `supporting_documents/background/` should be marked with footer-style endnotes or a clearly labeled source endnote block when page footers are not technically available.

Create retained UTF-8 text drafts under `outputs/drafts/`, render them to DOCX with `scripts/write_docx.py`, then render the DOCX files to PDF with `scripts/write_pdf.py --backend word`. If Word COM fails, fix Word COM or ask before using a fallback because PyMuPDF output will not fully match DOCX layout. Keep the text drafts, DOCX drafts, and PDF files:

```powershell
.\scripts\run_python.ps1 scripts/write_docx.py outputs/drafts/案例正文_ExampleCo.txt outputs/案例正文_ExampleCo.docx
.\scripts\run_python.ps1 scripts/write_docx.py outputs/drafts/案例使用说明_ExampleCo.txt outputs/案例使用说明_ExampleCo.docx
.\scripts\run_python.ps1 scripts/write_pdf.py outputs/案例正文_ExampleCo.docx outputs/案例正文_ExampleCo.pdf --backend word
.\scripts\run_python.ps1 scripts/write_pdf.py outputs/案例使用说明_ExampleCo.docx outputs/案例使用说明_ExampleCo.pdf --backend word
```

If Word COM and LibreOffice are unavailable, direct text-to-PDF rendering can be used only as an emergency fallback. Tell the user first that the PDF will not fully match the DOCX layout:

```powershell
.\scripts\run_python.ps1 scripts/write_pdf.py outputs/drafts/案例正文_ExampleCo.txt outputs/案例正文_ExampleCo.pdf
.\scripts\run_python.ps1 scripts/write_pdf.py outputs/drafts/案例使用说明_ExampleCo.txt outputs/案例使用说明_ExampleCo.pdf
```

To require Microsoft Word conversion and fail if Word or `pywin32` is unavailable, pass `--backend word`. To require LibreOffice conversion, pass `--backend libreoffice`.

Generated files inside `outputs/` are ignored by Git except for `.gitkeep`.

## Publishing To GitHub

Before pushing:

```powershell
git status --ignored
```

Confirm that the intended files are staged. `supporting_documents/examples/` and `supporting_documents/methodology/` are ignored by default because they are user-collected source materials; runtime templates live under `.codex_skills/case-writer/references/format-templates/`.
