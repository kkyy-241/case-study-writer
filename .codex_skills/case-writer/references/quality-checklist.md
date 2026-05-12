# Quality Checklist

## Before Drafting

- Ask the user to confirm the core content of this case: protagonist/company, financing event, decision conflict, course, and teaching emphasis.
- Ask the user to confirm the exact `<CompanyName>` suffix for retained drafts, final deliverables, and archive folders. If not provided, propose one and wait for confirmation; do not infer silently and do not mention any prior case company name.
- Before a new case-writing run, ask whether to archive current `supporting_documents/background/` and `outputs/` files into `outputs/archive/<CompanyName>/`, using the same company suffix as the deliverable filenames.
- Present a concise plan for approval after source preparation and before writing any draft prose, draft files, or final PDFs.
- Do not begin the student case, teaching note, narrative sections, or final output generation until the user explicitly approves the plan.
- Confirm that temporary extracted products are written under `outputs/text/`, not `artifacts/`.
- Confirm that retained build drafts are written under `outputs/drafts/` and kept after delivery.
- Confirm that final `.docx` and `.pdf` deliverables are retained in `outputs/` and use the company suffix, such as `案例正文_<CompanyName>.docx`.
- Confirm generated DOCX/PDF page headers show `商学院教学案例库` left-aligned and the confirmed `<CompanyName>` suffix right-aligned.
- Inspect `references/format-templates/安德科铭正文案例.pdf` and `references/format-templates/安德科铭案例使用说明.pdf` as the required format templates.
- If the background folder does not contain the confirmed core content, search public web sources or ask the user for supplementary materials.
- Search the web for the latest important public news about the entrepreneur, startup/company, investors, financing, products, regulation, litigation, IPO/M&A, bankruptcy, or other material events.
- Run `scripts/collect_images.py` when visuals would improve the case. If background materials lack suitable images, use web/image search to find appropriate public images, download selected direct image URLs through `collect_images.py`, and preserve source notes in the image manifest.
- Confirm that no charts, diagrams, or synthetic images are created by the workflow itself; visuals must come only from `background` extraction or web/image search.
- Identify protagonist, company, decision date, decision location, and decision deadline.
- Create a timeline of major operating and financing events.
- Create a financing facts table.
- Create an equity/control facts table if ownership matters.
- List missing facts and decide whether to state assumptions or omit analysis.
- Identify the central contradiction in one sentence.

## Student Case Checklist

- Is written in Chinese unless the user explicitly requested another language.
- Has a student-facing body length of 8,000-10,000 Chinese characters unless the user explicitly changed the target.
- Opens with a decision scene rather than a broad industry introduction.
- Uses real facts, dates, numbers, and named actors from background materials.
- Incorporates relevant latest-news facts when they materially affect the case context or decision pressure.
- Explains industry/product details only insofar as they affect finance, risk, valuation, or bargaining.
- Includes financing need and causes, not just financing events.
- Shows alternatives with genuine tradeoffs.
- Includes valuation, dilution, equity structure, control, investor terms, or financing instruments when relevant.
- Contains enough data for students to calculate or reason.
- Includes relevant images, charts, or exhibits when they clarify product context, financing tools, equity/control structure, market setting, or decision alternatives.
- Uses only extracted or web-searched images, not self-created diagrams or synthetic visuals.
- Does not include reference answers, theory lectures, or author judgment.
- Ends at an open decision node.
- Does not close with a resolved outcome, author recommendation, or teaching conclusion.
- Includes five logically progressive finance-focused discussion questions.
- Marks `background` facts with numbered note markers and includes footer-style source notes or a clearly labeled endnote block.
- Preserves source traceability for any images used as evidence or exhibits.

## Teaching Note Checklist

- Is written in Chinese unless the user explicitly requested another language.
- Includes case synopsis, teaching objectives, suitable courses and learners.
- Repeats the exact five discussion questions from the student case.
- Each answer introduces relevant theory and then applies it to case facts.
- Uses finance and entrepreneurial finance concepts rather than generic management commentary.
- Compares options where the case presents alternatives.
- Names assumptions when calculations or conclusions depend on incomplete data.
- Provides classroom-useful conclusions without pretending there is only one possible answer.
- Marks `background` facts with numbered note markers and includes footer-style source notes or a clearly labeled endnote block.

## Red Flags To Fix

- The case reads like company publicity or a chronological encyclopedia.
- The protagonist is unclear.
- The decision point appears only at the end.
- Financing terms are described but not connected to incentives, risk, valuation, dilution, or control.
- Questions can be answered without finance knowledge.
- The teaching note repeats the case story instead of analyzing it.
- The student case reveals the recommended decision or final real-world outcome too early.
