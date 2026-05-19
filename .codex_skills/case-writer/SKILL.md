---
name: case-writer
description: "Write finance-focused business school teaching cases and teaching notes from case background materials. Use when Codex needs to synthesize a background folder into independent Chinese teaching-case DOCX drafts and PDF files: a student-facing narrative case and an instructor-facing case teaching note, especially for undergraduate, master's, MBA, entrepreneurship finance, corporate finance, venture capital, valuation, equity structure, financing decisions, control rights, or investor negotiation contexts."
---

# Case Writer

## Overview

Use this skill to turn raw background materials into a business school teaching case package for finance and entrepreneurship finance classes. Produce two separate Chinese-language DOCX drafts and matching PDF files: a student-facing case narrative and an instructor-facing case teaching note.

This skill was created with Codex's built-in `skill-creator` workflow from the repository's collected case examples in `supporting_documents/examples/` and case-writing methodology materials in `supporting_documents/methodology/`. Repository users may update those folders with their own representative cases and writing-method requirements, then use `skill-creator` to revise this skill so its workflow and references match their preferred standards.

## Workflow

### Phase 0: Plan First And Consent Gate

Before any substantive drafting, ask the user what they want this case to emphasize. At minimum, ask for the desired protagonist or decision maker, company or event scope, core problem to highlight, course/audience, teaching emphasis, whether the case should be prospective or retrospective, and the exact `<CompanyName>` suffix to use for retained drafts, final deliverables, and archive folders.

The `<CompanyName>` suffix must be explicitly confirmed by the user during the plan stage. If the user has not provided it, propose one based on the case subject and ask for confirmation. Do not infer or silently choose the suffix from filenames, background materials, or prior case examples. Do not mention any prior case company name when asking for the suffix.

Every time the user asks to write a new case, confirm the `<CompanyName>` suffix and ask whether to archive the current `supporting_documents/background/` and `outputs/` files into `outputs/archive/<CompanyName>/` before drafting. The archive folder name must match the company suffix that will be used in the generated deliverable filenames. For example, if the user confirms `<CompanyName>` as `ExampleCo`, files such as `案例正文_ExampleCo.docx` correspond to `outputs/archive/ExampleCo/`. Use `scripts/archive_case.py <CompanyName> --mode copy` after the user approves archiving; use `--mode move` only when the user explicitly asks to clear the active folders.

Generated DOCX/PDF page headers must use `商学院教学案例库` left-aligned and the confirmed `<CompanyName>` suffix right-aligned. The right header text must match the filename suffix and archive folder name.

If the user has already provided some of those details, summarize them and ask for the missing details. Do not infer the case focus from file names alone.

After initial fact gathering, present a concise writing plan for explicit approval. The plan must include: proposed working title, confirmed `<CompanyName>` suffix, protagonist, decision date or decision window, central conflict, alternatives, expected finance theories, expected deliverables, archive plan, and source plan.

Wait for the user's clear approval before writing any student-facing case draft, instructor teaching note, final PDF, or substantial narrative section. Clear approval means the user says to proceed, start writing, begin drafting, generate the case, or otherwise unambiguously authorizes writing.

Before approval, only do preparatory work: dependency checks, source extraction into `outputs/`, source inventory, short fact-base notes, questions for clarification, and a proposed plan. Do not create draft case files, final output files, or narrative prose meant to be reused as the case.

### Phase 1: Source Preparation And Approved Drafting

1. If working inside this repository, use the repository wrapper `.\scripts\run_python.ps1` so the configured Conda environment is used instead of any shell-level Windows Store Python launcher. Run `.\scripts\run_python.ps1 scripts/check_dependencies.py` before extracting documents, generating DOCX drafts, or validating YAML. Required packages are `PyMuPDF` (`fitz`), `PyYAML` (`yaml`), and `python-docx` (`docx`).
2. Use the same wrapper to run `.\scripts\run_python.ps1 scripts/extract_text.py <background-folder> --output outputs/text/<name>` and create intermediate text products from PDF, DOCX, text, Markdown, or HTML background materials. Keep all intermediate products under `outputs/`; do not use `artifacts/` for case-writing runs.
3. Before drafting, ask the user to state the core focus of the case for this run, such as the target entrepreneur, startup/company, financing event, decision conflict, course, and desired teaching emphasis. If the background materials do not contain that core content, use web search to fill public facts or ask the user to provide more materials when the missing content is not publicly available.
4. Read the user's background folder before drafting. Build a fact base with dates, actors, company milestones, financing events, amounts, valuations, ownership terms, investor options, conflicts, market data, and sources.
5. Use web search before writing the student case to check for the latest important news about the entrepreneur, startup/company, financing round, investors, products, regulation, litigation, IPO/M&A, bankruptcy, or other material events. Incorporate relevant, verified, up-to-date news into the case facts with source traceability. If web access is unavailable, state that limitation and ask the user whether to proceed from local materials only.
6. Identify the case core: protagonist, decision deadline, financing dilemma, viable alternatives, missing information, and the finance theories students should apply.
7. Before drafting, inspect `references/format-templates/安德科铭正文案例.pdf` and `references/format-templates/安德科铭案例使用说明.pdf` as the required format templates. Match their document style, section logic, teaching-case tone, and instructor-note organization unless the user requests a different template.
8. Read only the references needed for the task:
   - `references/case-writing-method.md` for case method principles and narrative standards.
   - `references/student-case-structure.md` before writing the student-facing case.
   - `references/teaching-note-structure.md` before writing the instructor note.
   - `references/entrepreneurial-finance-frameworks.md` when designing questions and reference answers.
   - `references/quality-checklist.md` before final delivery.
9. Draft the student case first as a retained UTF-8 text source in `outputs/drafts/案例正文_<CompanyName>.txt`, then generate `outputs/案例正文_<CompanyName>.docx` with `.\scripts\run_python.ps1 scripts/write_docx.py outputs/drafts/案例正文_<CompanyName>.txt outputs/案例正文_<CompanyName>.docx`. Keep the case story-like, factual, decision-centered, and free of instructor analysis or final answers. The student-facing case body should be 8,000-10,000 Chinese characters unless the user changes the target length.
10. End the student-facing case with an open decision node. The ending must leave the protagonist facing unresolved alternatives and must not reveal the recommended answer.
11. Draft five discussion questions at the end of the student case. Make them logically progressive from fact diagnosis to financial analysis to decision judgment.
12. Draft the teaching note second as a retained UTF-8 text source in `outputs/drafts/案例使用说明_<CompanyName>.txt`, then generate `outputs/案例使用说明_<CompanyName>.docx` with `.\scripts\run_python.ps1 scripts/write_docx.py outputs/drafts/案例使用说明_<CompanyName>.txt outputs/案例使用说明_<CompanyName>.docx`. Reuse the same five questions and provide theory-grounded reference answers tied to case facts. The teaching note should be 10,000-15,000 Chinese characters unless the user changes the target length.
13. Convert PDF deliverables from the DOCX files, not directly from text drafts. For final deliverables, use Microsoft Word COM explicitly so PDF layout matches DOCX: `.\scripts\run_python.ps1 scripts/write_pdf.py outputs/案例正文_<CompanyName>.docx outputs/案例正文_<CompanyName>.pdf --backend word` and `.\scripts\run_python.ps1 scripts/write_pdf.py outputs/案例使用说明_<CompanyName>.docx outputs/案例使用说明_<CompanyName>.pdf --backend word`. If Word COM fails, stop and tell the user instead of silently falling back to PyMuPDF. Keep retained text drafts in `outputs/drafts/` and both DOCX drafts and PDFs in `outputs/`.
14. At the end of the run, delete non-deliverable intermediate products outside the retained `.txt`, `.docx`, and `.pdf` deliverables, including extracted text folders under `outputs/text/` unless the user explicitly asks to keep them.
15. Mark information sources from the `background` folder as footer-style endnotes. Use numbered note markers in the body and place concise source notes at the bottom of the relevant page, or in a clearly labeled `资料来源尾注` block if true per-page footers are not technically available in the draft format.

### Image Asset Workflow

When visuals would help the case, use repository scripts for image preparation. First run `.\scripts\run_python.ps1 scripts/collect_images.py <background-folder> --output outputs/drafts/images` to extract candidate images from PDF/DOCX/background image files and create `manifest.json` and `manifest.md`. If the background materials do not contain suitable images for exhibits, product/company context, financing diagrams, equity/control visuals, or market screenshots, use web/image search to identify appropriate public images, then download only selected direct image URLs with `scripts/collect_images.py --url <image-url> --source-note <title/publisher/date/source-page> --output outputs/drafts/images`.

Draft sources may insert collected images with Markdown image syntax, such as `![图1 股权结构示意图](images/ownership-chart.png)` from files stored under `outputs/drafts/images/`. `scripts/write_docx.py` embeds these images in DOCX. Final PDFs should still be generated from DOCX through Word COM or LibreOffice whenever possible. Direct `txt -> PDF` rendering is only an emergency fallback after telling the user that its layout will not fully match the DOCX. Preserve image source traceability in the image manifest and in the case source notes when a visual is used as evidence.

Do not create charts, diagrams, or synthetic images yourself. The workflow may only use images extracted from `supporting_documents/background/` or images found through web/image search. This keeps image formats and provenance stable.

## Non-Negotiables

- Use `scripts/collect_images.py` for local image extraction and web-sourced image downloads. If background materials lack suitable images and visuals would improve the case, use web/image search to identify relevant public images, download only selected direct image URLs, and keep title/publisher/date/source-page traceability.
- Never create charts, diagrams, or synthetic images yourself. Use only images from `supporting_documents/background/` or from web/image search.
- Before starting any new case-writing run, ask whether to archive the current background and output files into `outputs/archive/<CompanyName>/`, where `<CompanyName>` exactly matches the deliverable filename suffix.
- Write both final PDF files in Chinese. Only switch languages if the user explicitly requests another language.
- Always run the plan-first consent gate before drafting. Ask for the run-specific case focus, propose a plan, and wait for explicit user approval before writing draft prose or final outputs.
- Do not use `artifacts/` for case-writing intermediate products. Put temporary extracted text under `outputs/text/`, retained build drafts under `outputs/drafts/`, and generated DOCX/PDF deliverables under `outputs/`. After cleanup, keep the retained `.txt` build drafts plus final `.docx` and `.pdf` deliverables.
- Use `references/format-templates/安德科铭正文案例.pdf` and `references/format-templates/安德科铭案例使用说明.pdf` as the default formatting and organization templates.
- The student-facing case body must be 8,000-10,000 Chinese characters unless the user explicitly changes the length.
- The teaching note must be 10,000-15,000 Chinese characters unless the user explicitly changes the length.
- Do not invent facts, amounts, dates, investor names, financing terms, or outcomes. If a needed fact is absent, state an assumption or mark it as unavailable.
- Check the latest important public news by web search before writing the student-facing case, and cite or otherwise trace the sources used.
- Do not reveal the real outcome after the decision point in the student-facing case unless the background materials and user explicitly require a retrospective case.
- Keep the student case analytical enough for finance work but not answer-like. Put theories, teaching logic, and recommended decisions in the teaching note.
- `案例正文`不得以教师或学生为隐含读者；所有教学建议必须写入`案例使用说明`。
- 在复审时，正文不得使用“对学生而言”、“学生可能会”、“教师应当”等隐含读者的表述；正文应保持客观叙述，教学建议仅存在于教学说明。
- Preserve source traceability for important facts. Add source notes or citations when the background materials provide them.
- For facts from `background`, use footer-style endnotes with source file/title, publisher or website when available, and date or access date when available.
- If background materials conflict, flag the conflict and choose the version with the stronger source, or present the uncertainty as part of the case.
- The student-facing case must close with an open-ended decision point for classroom discussion.

## Output Contract

Current repository output rules supersede any older wording in this file: create retained text drafts under `outputs/drafts/`, create DOCX drafts with `scripts/write_docx.py`, keep the DOCX drafts, and convert DOCX files to PDFs with `scripts/write_pdf.py --backend word`. For final style fidelity, use DOCX input and Word COM; do not use direct text-to-PDF generation for final deliverables unless the user explicitly accepts layout differences. The retained deliverables are six files: `outputs/drafts/案例正文_<CompanyName>.txt`, `outputs/drafts/案例使用说明_<CompanyName>.txt`, `outputs/案例正文_<CompanyName>.docx`, `outputs/案例正文_<CompanyName>.pdf`, `outputs/案例使用说明_<CompanyName>.docx`, and `outputs/案例使用说明_<CompanyName>.pdf`.

Create the following substantive Chinese-language deliverables unless the user asks for more:

1. `案例正文`: for student pre-class reading. Include title, protagonist and company background, industry and market context, product/service and business model, entrepreneurial development, latest important news, financing needs and causes, key financing decisions, valuation disputes, control-right conflicts or investor bargaining, financing instruments and equity structure, important facts/data/exhibits, five discussion questions, open decision ending, and footer-style endnotes for `background` sources.
2. `案例使用说明`: for instructors. Include case synopsis, teaching objectives, suitable courses and learners, and reference answers to the five discussion questions. Each answer should introduce relevant finance or entrepreneurial finance theory, apply it to the case, compare decision options when appropriate, and preserve source traceability for `background` facts with footer-style endnotes.

## Style

Use concrete scenes, dates, people, dialogue, boardroom or negotiation tension, and operational details in the student case. Use academic and practical logic in the teaching note: theory first, case application second, decision implications third.

Avoid promotional writing, generic management advice, moralizing, unsupported adjectives, and omniscient conclusions. The case should make students work.

For this repository, use `references/format-templates/安德科铭正文案例.pdf` and `references/format-templates/安德科铭案例使用说明.pdf` as the default format templates. The student-facing case body should be 8,000-10,000 Chinese characters unless the user explicitly changes the requirement. The teaching note should be 10,000-15,000 Chinese characters unless the user changes the target length.
