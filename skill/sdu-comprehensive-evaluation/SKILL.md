---
name: sdu-comprehensive-evaluation
description: Project-local skill for organizing, auditing, and summarizing Shandong University Low Altitude Science and Engineering College undergraduate comprehensive evaluation materials. Use when working in this project on 综测/综合评价 materials, student submission folders, proof documents, naming cleanup, scoring table checks, Excel .xlsx summaries, Word .docx proof packs, PDF rules, or 山东大学低空科学与工程学院素质能力评价.
---

# SDU Comprehensive Evaluation

## Scope

Use this skill in a workspace that contains the official `综测模板` materials, or in the published skill repository after copying `project-files/综测模板` into the active comprehensive-evaluation workspace.

The project contains the authoritative notice, scoring rules, competition classification files, Excel template, and example submission. Do not invent scoring rules. If a case is ambiguous, mark it for manual review and cite the source file that controls the decision.

## Project File Map

Use these local files deliberately:

- `山东大学低空科学与工程学院本科生综合评价办法(试行).pdf`: main college comprehensive evaluation method and detailed score caps.
- `综测模板/关于开展2025-2026学年综合评价工作的通知.docx`: collection period, deadline, submission/naming requirements, and current-year reminders.
- `综测模板/关于《山东大学本科生综合评价办法（试行）》创新素养部分的补充通知（2025）.docx`: latest local innovation-quality supplement; prefer it for innovation items when it is more specific or newer than the PDF.
- `综测模板/附件1：山东大学本科生学科双创竞赛分级分类（2024试行版）.docx`: competition classification reference for A, A-, B, C, D, E.
- `综测模板/附件2：学院申报学科专业竞赛备案名单（2024试行版）.docx`: college-filed A*, A-*, B* reference. Do not assume another college's filing applies to Low Altitude Science and Engineering.
- `综测模板/【示例+注意事项】低空科学与工程学院2025级XX班25-26学年本科生素质能力评价成绩.xlsx`: official Excel template and cell-level filling notes.
- `综测模板/低空科学与工程学院拓展培养学分置换方案.docx`: related background only. In this project, treat拓培系统材料 as corresponding to志愿学时 evidence; do not use it as a replacement for综测 scoring rules.
- `综测模板/张三（提交示例！）/xxx综测统计.xlsx`: example filled-style statistics workbook.
- `综测模板/张三（提交示例！）/张三综测证明材料.docx`: example proof pack shell with four category headings.

## End-to-End Student Package Flow

1. Locate the target material folder and preserve originals. If cleanup is requested, create normalized copies or clearly scoped edits instead of destructive renames unless the user asks for renaming.
2. Read `references/rules.md` before scoring or judging eligibility. For disputes, reopen the original project document listed in that reference.
3. Ask for facts, not scores. Collect class, student ID, moral grade, fitness/psychology participation, volunteer hours, social-practice count, dorm status, roles, and whether each proof is individual/team plus organizer/level/rank/order. Then calculate scores from rules.
4. Classify proof materials into the four official categories and their基础性/成果性 subfolders first. Mark ambiguous items for manual confirmation, and only rename/move after showing a planned move table. When the user gives an explicit local口径 for an ambiguous item, use that decision consistently in the folder names, Excel score reasons, and Word proof pack.
5. Generate or update the Excel from the official template as the score ledger for the same organized proof set, not as a separate scratch artifact. The template has a merged sample block at `A7:E8`; for a one-student workbook, unmerge/remove the sample block, put the official student row at row 7, delete the example/notes rows below, and verify `A7:E7` contains序号、专业班级、学号、姓名、德育考核等级.
6. Rename final category folders with category totals, then rename subcategory folders with sub-scores following the same scores recorded in Excel, such as `张三文艺素养  12分/张三文艺素养成果性评价证明材料  3分`.
7. Treat a root-level proof pack `.docx`, such as `姓名综测证明材料.docx`, as a live deliverable. Whenever adding, deleting, reclassifying, or renaming proof materials, update or regenerate the Word proof pack from the organized folders and the same Excel row before finalizing. When asked to "看模板/整理证明材料/把图片放到对应栏目", use `综测模板/张三（提交示例！）/张三综测证明材料.docx` and run:

```bash
python3 <skill-folder>/scripts/build_proof_pack.py <submission-folder>
```

8. Render the generated `.docx` with the documents skill `render_docx.py` and visually inspect page PNGs before delivery. Fix stale template headers, blank pages, image overflow, and headings separated from images before reporting completion.
9. Run the submission inspector as a final package check after folder organization, Excel generation, and Word proof-pack generation:

```bash
python3 <skill-folder>/scripts/inspect_submission.py <submission-folder>
```

10. For each student, check six things in order: folder structure, proof filenames, Excel official row/basic information, score reasons, score caps, and whether the Word proof pack reflects the same folder categories and proof filenames.
11. Produce concise Chinese output: issues first, then suggested fixes, then any items requiring manual confirmation.

## Expected Student Folder

Use the example under `综测模板/张三（提交示例！）/` as the structure model.

Expected contents:

- One student folder named by student name.
- One comprehensive evaluation statistics `.xlsx` file.
- Optional one root-level proof pack `.docx`, such as `张三综测证明材料.docx`. If present, it must be kept synchronized with the final proof folders and Excel; do not treat it as a disposable example file.
- Proof materials grouped under four category folders containing `身心素养`, `文艺素养`, `劳动素养`, `创新素养`.
- For final organized submissions, category folders must prefix the student name and suffix the category total score, following the example style such as `张三文艺素养  12分`.
- Inside each category, create `基础性评价` and `成果性评价` proof folders with sub-scores, following the example style such as `张三文艺素养成果性评价证明材料  3分`. Empty `0分` folders are acceptable when a category/subcategory has no proof but the final folder structure is being prepared.
- Electronic proof filenames should be understandable and include the item plus score, such as `xx比赛 xx分`. For innovation contest items, include contest category, level, award/rank, and personal order.

## Word Proof Pack Rules

Use the example Word shell:

`综测模板/张三（提交示例！）/张三综测证明材料.docx`

For each student's final proof pack:

- Output to the student folder as `<姓名>综测证明材料.docx`; do not modify the example template.
- Update or regenerate this proof pack whenever proof files or Excel score reasons change. This includes newly found awards, deleted `0分` materials, reclassified items, renamed proof files, and score changes.
- Read class, student ID, and name from the official Excel row `B7:D7` when available, and replace the template header with `<专业班级> <学号> <姓名> 综测证明材料`.
- Insert proof pictures by folder category, in this order: `身心素养`, `文艺素养`, `劳动素养`, `创新素养`.
- Within each category, insert `基础性评价` pictures before `成果性评价` pictures. Use the proof filename stem as the centered caption above each image.
- For each counted proof, include a concise title with the final score, then the proof image/document content when practical. If a counted PDF is an auxiliary proof and the award certificate image is already included, at minimum list the PDF title in the same section so the pack does not look stale.
- If a material is removed from the final submission because it does not count, remove it from the Word proof pack too.
- Put each proof image on its own page using paragraph `page_break_before`, not a standalone manual page-break paragraph; this avoids blank pages when an image nearly fills the previous page.
- Scale images proportionally to fit within the page content width and leave room for category/subcategory/caption text. Do not crop proof images.
- If a category has no pictures, keep a category page that says `本栏目暂无图片证明材料。`.
- After generation or editing, render and inspect the DOCX. Confirm every picture appears under the correct category, no page is blank unintentionally, no image is clipped, and the header no longer contains the sample `张三` information. Clean render intermediates unless the user asks to keep them.

## Excel Editing Rules

Use the project Excel template:

`综测模板/【示例+注意事项】低空科学与工程学院2025级XX班25-26学年本科生素质能力评价成绩.xlsx`

Important columns:

- F/G: 身心素养基础性评价 reason/score, cap 9.
- H/I: 身心素养成果性评价 reason/score, cap 6.
- J/K: 文艺素养基础性评价 reason/score, cap 9.
- L/M: 文艺素养成果性评价 reason/score, cap 6.
- N/O: 劳动素养基础性评价 reason/score, cap 15.
- P/Q: 劳动素养成果性评价 reason/score, cap 10.
- R/S: 创新素养基础性评价 reason/score, cap 5.
- T/U: 创新素养成果性评价 reason/score, cap 40.
- V: total score, cap 100.
- W: remarks.

For score reasons, keep one numbered item per line. Include full award/activity name, level, award/rank, individual/team status, personal order such as `(3/8)` when relevant, and explicit `+x分`. If an item is listed but not counted, keep it and write `+0分` with a brief reason.

Final workbook guardrails:

- Do not write student data into row 8 while `A7:E8` is still merged as the template sample block; this hides basic information. The official one-student row should be row 7 after removing the sample/notes area.
- Delete template-only sample and注意事项 rows from final student workbooks unless the user explicitly wants a workbook retaining examples.
- Keep final Excel text submission-ready. Remove internal phrases such as `按人工口径`, `待人工确认`, `需确认`, `不计创新`, and long explanatory notes once the user has approved the local decision.
- If a material does not count and the user asks for a direct submission package, remove it from both Excel and proof folders instead of keeping a `+0分` proof, unless the user explicitly wants a review archive.
- Verify `V7` (or the active student row) sums `G/I/K/M/O/Q/S/U` and remains within the 100-point cap.

## Scoring Guardrails

Load `references/rules.md` for detailed caps and special cases.

Core caps:

- 素质能力总分: 100 = 身心 15 + 文艺 15 + 劳动 25 + 创新 45.
- 身心: 基础 9, 成果 6.
- 文艺: 基础 9, 成果 6.
- 劳动: 基础 15, 成果 10.
- 创新: 基础 5, 突破提升/成果 40.

General principles:

- Same project in the same indicator counts only the highest applicable score.
- Proof materials should have official stamp or credible official proof when required.
- The 2025-2026 collection notice covers materials from 2025-09-01 to 2026-06-30 for summer collection, with 2026-07-01 to 2026-08-31 collected later after school starts.
-拓培系统 records correspond to志愿学时 in this project. Use them to verify劳动素养基础性评价中的志愿服务 hours, then apply the comprehensive evaluation rule: 20 volunteer hours = 3 points, 30 volunteer hours = 5 points.
- Innovation breakthrough materials are collected after school starts; do not mix them into the summer submission unless the user explicitly asks to prepare a later-review list.
- If later official files conflict with local references, follow the latest official file and report the conflict.

## Output Patterns

For auditing one folder:

- `通过`: list confirmed complete parts.
- `需修改`: list concrete filename/table/folder fixes.
- `待人工确认`: list ambiguous scoring or eligibility questions with source rule.

For reorganizing:

- Show the planned rename/move table before changing files.
- Keep original files available.
- Avoid deleting duplicates unless the user explicitly confirms.
