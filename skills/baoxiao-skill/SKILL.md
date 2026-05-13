---
name: baoxiao-skill
description: Operate HuiLianYi personal reimbursement and public-payment workflows for Chinese VAT invoices, no-ticket expenses, payment screenshots, fixed-asset matching, Excel reimbursement workbooks, reimbursement overview tables, report-number writeback, attachment submission lists, and A5 invoice printing. Use when Codex needs to distinguish personal reimbursement from public payment, match invoices or no-ticket screenshots to fixed-asset rows, create or update HuiLianYi reports, upload attachments, submit or hold drafts, create public payment orders, or print invoices and payment evidence.
---

# HuiLianYi Reimbursement

Use this skill for HuiLianYi expense work that mixes browser operations, invoices, screenshots, and Excel ledgers.

## Core Rules

- Use Chrome automation for HuiLianYi because it depends on the user's logged-in session.
- Use spreadsheet tooling for `.xlsx` edits and verify totals after writing.
- Extract PDF invoice text with bundled Python `pypdf` before deciding invoice type, buyer, seller, amount, and bank account details.
- Preserve user files. Write backups before modifying workbooks or regenerating submission lists.
- Do not click final submit unless the user explicitly asks to submit. If the user says "先不要最终提交", save drafts only.
- Decide whether the work is `个人报销` or `对公付款` before creating any HuiLianYi document. If the payer or recipient is unclear, stop and ask.
- HuiLianYi titles must use Chinese Han characters, not pinyin. Use Chinese purpose words or the official Chinese party short name; do not create titles with romanized Chinese.
- When fixed-asset or equipment rows are present, match invoice PDFs and no-ticket screenshots back to `资产编号`, `使用人`, `使用部门`, amount, and purchase reason before building the reimbursement workbook.
- When rebuilding reports, treat previous validated workbooks, existing HuiLianYi drafts/reports, screenshots, WeWork cache attachments, and the Desktop invoice folder as separate input sources. Reconcile totals before entering the browser; do not rely on only the newest workbook if an older report contains a larger amount.
- Editable HuiLianYi text fields must not contain the sensitive word `无票`. The system document type may display no-ticket wording, but titles,事由说明, remarks, uploaded list names, and user-facing workbook labels should use neutral Chinese wording such as `凭证费用`, `支付凭证费用`, or `票据缺失类费用`.
- A request like `重新建一套` means create a genuinely new HuiLianYi document set. Do not reuse an old zero-amount draft unless the user explicitly approves reuse.
- If a required field cannot be inferred safely, stop and ask. Common blockers: fixed asset code, missing supplier, missing bank account, unknown recipient account, or mismatched screenshot amount.

## Standard Workflow

1. Gather inputs:
   - Source workbook and target output workbook.
   - Invoice folder on Desktop, direct WeWork cache files, and payment screenshots.
   - User intent: submit now, save draft, print, or write back IDs.
2. Classify document flow:
   - `个人报销`: the employee/login user paid first and the company reimburses them. Payment recipient should be the logged-in applicant.
   - `对公付款`: the company pays an external party. Payment recipient should be the invoice issuer or contract counterparty, not the logged-in applicant.
   - When invoice issuer and contract counterparty differ, ask which party should receive payment before continuing.
3. Classify invoices:
   - `增值税专用发票` means VAT special invoice.
   - `普通发票` means ordinary invoice.
   - Foreign invoices or payment screenshots are not Chinese VAT invoices; treat them as no-ticket evidence unless the user says otherwise.
   - Group reimbursement data by `发票抬头`; same title may be one report or one fee line when the user allows it.
4. Build or update Excel:
   - Maintain a detail sheet with at least: `单据类型`, `收款对象`, `发票抬头`, `费用名称`, `资产编号`, `使用人`, `使用部门`, `金额`, `采购原因`, report id/type/amount, and notes.
   - Generate the submission list in this order: `发票抬头`, `费用名称`, `资产编号`, `使用人`, `使用部门`, `金额`, `采购原因`.
   - Generate `报销一览表` in this order: `发票抬头`, `费用名称`, `资产编号`, `使用人`, `使用部门`, `金额`, `汇联易报销单号`, `采购原因`, `报销单类型`, `汇联易单据金额`, `报销备注`.
   - Write HuiLianYi report ids back to every relevant detail row and summary row.
   - Before browser entry, compare the rebuilt totals against prior validated workbooks or existing HuiLianYi report totals. If a prior amount is larger, explain the delta by item and stop until the missing rows are added or the user confirms they should be excluded.
5. Operate HuiLianYi:
   - For payment-proof rows without VAT invoices, choose the no-ticket/hand-entered fee flow, do not attach invoices as invoice records, and upload payment proof plus the latest submission list as attachments.
   - For invoice rows, use invoice generation or invoice upload. Use direct file upload; do not import zip packages unless the user asks.
   - For personal reimbursement, verify the payee is the logged-in applicant.
   - For public payment, verify the payee is the invoice issuer or contract counterparty.
   - After saving, verify visible report amount, line count, expense numbers, and status text.
6. Print when requested:
   - Use A5 paper and fit-to-page.
   - VAT special invoices: 2 copies.
   - Ordinary invoices: 1 copy.
   - Payment screenshots/images: 1 copy.
   - Avoid duplicate printing from both the original invoice folder and a classified-copy folder.

## HuiLianYi Browser Notes

- The app commonly shows a missing-invoice warning for no-ticket fees. Use `继续保存` only when the user requested no-ticket handling.
- If a fee type such as `电子设备(3年)` requires `固定资产编码`, ask for the code. If the user says no fixed asset code is needed, switch to an appropriate low-value consumable type.
- Personal reimbursement titles should use a Chinese format such as `YYYY年MM月办公费用报销` or `YYYY年MM月软件会员费用报销`.
- For payment-proof/no-ticket reports, use a Chinese title such as `YYYY年MM月凭证费用报销` or `YYYY年MM月支付凭证费用报销`. Do not use pinyin, and do not put `无票` in editable fields even if HuiLianYi's fixed system type displays it.
- Record important HuiLianYi identifiers in the final answer and workbook: report number, expense number, amount, and whether the report was submitted or left editing.

## Public Payment Orders

For `对公支付单`:

- Title format should use Chinese Han characters and follow HuiLianYi guidance: time + recipient/purpose + payment purpose, for example `YYYY年MM月供应商服务费对公付款`.
- Set `归属公司` to the invoice buyer and `往来公司(收款公司)` to the invoice seller.
- Prefer `到票支付` for VAT invoices and generate/import the fee from the invoice file.
- Verify seller bank account from the invoice before submission:
  - `销方开户银行`
  - `银行账号`
- If the supplier does not exist in HuiLianYi, ask before adding it.

## Bundled Resources

- `scripts/classify_and_print.py`: classify invoice PDFs/images and optionally send A5 print jobs with the standard copy rules.
- `scripts/build_reimbursement_overview.py`: match evidence rows to fixed-asset rows and generate a formatted `报销一览表` workbook; run it with a Python environment that has `openpyxl`, such as the bundled spreadsheet runtime.
- `scripts/reconcile_reimbursement_totals.py`: compare rebuilt reimbursement data against prior workbooks or HuiLianYi exports before browser entry; use it to catch missing payment screenshots or rows when totals differ.
- `references/huilianyi-field-guide.md`: field mapping, Excel writeback, and browser-operation details.
