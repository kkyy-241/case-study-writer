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

Before any substantive drafting, ask the user what they want this case to emphasize. At minimum, ask for the desired protagonist or decision maker, company or event scope, core problem to highlight, course/audience, teaching emphasis, and whether the case should be prospective or retrospective.

If the user has already provided some of those details, summarize them and ask for the missing details. Do not infer the case focus from file names alone.

After initial fact gathering, present a concise writing plan for explicit approval. The plan must include: proposed working title, protagonist, decision date or decision window, central conflict, alternatives, expected finance theories, expected deliverables, and source plan.

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
9. Draft the student case first as a non-deliverable UTF-8 text source in `outputs/text/drafts/案例正文.txt`, then generate `outputs/案例正文.docx` with `.\scripts\run_python.ps1 scripts/write_docx.py outputs/text/drafts/案例正文.txt outputs/案例正文.docx`. Keep the case story-like, factual, decision-centered, and free of instructor analysis or final answers. The student-facing case body should be 8,000-10,000 Chinese characters unless the user changes the target length.
10. End the student-facing case with an open decision node. The ending must leave the protagonist facing unresolved alternatives and must not reveal the recommended answer.
11. Draft five discussion questions at the end of the student case. Make them logically progressive from fact diagnosis to financial analysis to decision judgment.
12. Draft the teaching note second as a non-deliverable UTF-8 text source in `outputs/text/drafts/案例使用说明.txt`, then generate `outputs/案例使用说明.docx` with `.\scripts\run_python.ps1 scripts/write_docx.py outputs/text/drafts/案例使用说明.txt outputs/案例使用说明.docx`. Reuse the same five questions and provide theory-grounded reference answers tied to case facts. The teaching note should be 10,000-15,000 Chinese characters unless the user changes the target length.
13. Convert DOCX drafts with `.\scripts\run_python.ps1 scripts/write_pdf.py outputs/案例正文.docx outputs/案例正文.pdf` and `.\scripts\run_python.ps1 scripts/write_pdf.py outputs/案例使用说明.docx outputs/案例使用说明.pdf`. For DOCX inputs, `write_pdf.py` defaults to `--backend auto`: it tries Microsoft Word COM first on Windows, then LibreOffice, then a lightweight PyMuPDF text renderer. Keep both DOCX drafts and PDFs in `outputs/`.
14. At the end of the run, delete non-deliverable intermediate products outside the retained `.docx` and `.pdf` deliverables, including extracted text folders under `outputs/text/` unless the user explicitly asks to keep them.
15. Mark information sources from the `background` folder as footer-style endnotes. Use numbered note markers in the body and place concise source notes at the bottom of the relevant page, or in a clearly labeled `资料来源尾注` block if true per-page footers are not technically available in the draft format.

## Non-Negotiables

- Write both final PDF files in Chinese. Only switch languages if the user explicitly requests another language.
- Always run the plan-first consent gate before drafting. Ask for the run-specific case focus, propose a plan, and wait for explicit user approval before writing draft prose or final outputs.
- Do not use `artifacts/` for case-writing intermediate products. Put intermediate text, DOCX drafts, and generated PDFs under `outputs/`; keep only the final `.docx` and `.pdf` deliverables there after cleanup.
- Use `references/format-templates/安德科铭正文案例.pdf` and `references/format-templates/安德科铭案例使用说明.pdf` as the default formatting and organization templates.
- The student-facing case body must be 8,000-10,000 Chinese characters unless the user explicitly changes the length.
- The teaching note must be 10,000-15,000 Chinese characters unless the user explicitly changes the length.
- Do not invent facts, amounts, dates, investor names, financing terms, or outcomes. If a needed fact is absent, state an assumption or mark it as unavailable.
- Check the latest important public news by web search before writing the student-facing case, and cite or otherwise trace the sources used.
- Do not reveal the real outcome after the decision point in the student-facing case unless the background materials and user explicitly require a retrospective case.
- Keep the student case analytical enough for finance work but not answer-like. Put theories, teaching logic, and recommended decisions in the teaching note.
- Preserve source traceability for important facts. Add source notes or citations when the background materials provide them.
- For facts from `background`, use footer-style endnotes with source file/title, publisher or website when available, and date or access date when available.
- If background materials conflict, flag the conflict and choose the version with the stronger source, or present the uncertainty as part of the case.
- The student-facing case must close with an open-ended decision point for classroom discussion.

## Output Contract

Current repository output rules supersede any older wording in this file: create DOCX drafts with `scripts/write_docx.py`, keep the DOCX drafts, and convert them to PDFs with `scripts/write_pdf.py`. The PDF converter tries Microsoft Word COM first on Windows, then LibreOffice, then PyMuPDF text rendering when layout-preserving converters are unavailable. The retained deliverables are four files in `outputs/`: `案例正文.docx`, `案例正文.pdf`, `案例使用说明.docx`, and `案例使用说明.pdf`. Keep the DOCX drafts after PDF conversion and delete only non-deliverable intermediate text sources.

Create the following substantive Chinese-language deliverables unless the user asks for more:

1. `案例正文`: for student pre-class reading. Include title, protagonist and company background, industry and market context, product/service and business model, entrepreneurial development, latest important news, financing needs and causes, key financing decisions, valuation disputes, control-right conflicts or investor bargaining, financing instruments and equity structure, important facts/data/exhibits, five discussion questions, open decision ending, and footer-style endnotes for `background` sources.
2. `案例使用说明`: for instructors. Include case synopsis, teaching objectives, suitable courses and learners, and reference answers to the five discussion questions. Each answer should introduce relevant finance or entrepreneurial finance theory, apply it to the case, compare decision options when appropriate, and preserve source traceability for `background` facts with footer-style endnotes.

## Style

Use concrete scenes, dates, people, dialogue, boardroom or negotiation tension, and operational details in the student case. Use academic and practical logic in the teaching note: theory first, case application second, decision implications third.

Avoid promotional writing, generic management advice, moralizing, unsupported adjectives, and omniscient conclusions. The case should make students work.

For this repository, use `references/format-templates/安德科铭正文案例.pdf` and `references/format-templates/安德科铭案例使用说明.pdf` as the default format templates. The student-facing case body should be 8,000-10,000 Chinese characters unless the user explicitly changes the requirement. The teaching note should be 10,000-15,000 Chinese characters unless the user changes the target length.
