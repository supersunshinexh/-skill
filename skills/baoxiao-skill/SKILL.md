---
name: baoxiao-skill
description: Operate HuiLianYi reimbursement and public-payment workflows for Chinese VAT invoices, no-ticket expenses, payment screenshots, Excel reimbursement workbooks, report-number writeback, attachment submission lists, and A5 invoice printing. Use when Codex needs to classify invoices by title or VAT type, create or update HuiLianYi reports, upload attachments, submit or hold drafts, create public payment orders, or print invoices and payment evidence.
---

# HuiLianYi Reimbursement

Use this skill for HuiLianYi expense work that mixes browser operations, invoices, screenshots, and Excel ledgers.

## Core Rules

- Use Chrome automation for HuiLianYi because it depends on the user's logged-in session.
- Use spreadsheet tooling for `.xlsx` edits and verify totals after writing.
- Extract PDF invoice text with bundled Python `pypdf` before deciding invoice type, buyer, seller, amount, and bank account details.
- Preserve user files. Write backups before modifying workbooks or regenerating submission lists.
- Do not click final submit unless the user explicitly asks to submit. If the user says "先不要最终提交", save drafts only.
- If a required field cannot be inferred safely, stop and ask. Common blockers: fixed asset code, missing supplier, missing bank account, unknown recipient account, or mismatched screenshot amount.

## Standard Workflow

1. Gather inputs:
   - Source workbook and target output workbook.
   - Invoice folder on Desktop, direct WeWork cache files, and payment screenshots.
   - User intent: submit now, save draft, print, or write back IDs.
2. Classify invoices:
   - `增值税专用发票` means VAT special invoice.
   - `普通发票` means ordinary invoice.
   - Foreign invoices or payment screenshots are not Chinese VAT invoices; treat them as no-ticket evidence unless the user says otherwise.
   - Group reimbursement data by `发票抬头`; same title may be one report or one fee line when the user allows it.
3. Build or update Excel:
   - Maintain a detail sheet with at least: `发票抬头`, `费用名称`, `资产编号`, `使用人`, `使用部门`, `金额`, `采购原因`, report id/type/amount, and notes.
   - Generate the submission list in this order: `发票抬头`, `费用名称`, `资产编号`, `使用人`, `使用部门`, `金额`, `采购原因`.
   - Write HuiLianYi report ids back to every relevant detail row and summary row.
4. Operate HuiLianYi:
   - For no-ticket rows, choose `无票报销` / hand-entered fee, do not attach invoices as invoice records, and upload payment proof plus the latest submission list as attachments.
   - For invoice rows, use invoice generation or invoice upload. Use direct file upload; do not import zip packages unless the user asks.
   - After saving, verify visible report amount, line count, expense numbers, and status text.
5. Print when requested:
   - Use A5 paper and fit-to-page.
   - VAT special invoices: 2 copies.
   - Ordinary invoices: 1 copy.
   - Payment screenshots/images: 1 copy.
   - Avoid duplicate printing from both the original invoice folder and a classified-copy folder.

## HuiLianYi Browser Notes

- The app commonly shows a missing-invoice warning for no-ticket fees. Use `继续保存` only when the user requested no-ticket handling.
- If a fee type such as `电子设备(3年)` requires `固定资产编码`, ask for the code. If the user says no fixed asset code is needed, switch to an appropriate low-value consumable type.
- Record important HuiLianYi identifiers in the final answer and workbook: report number, expense number, amount, and whether the report was submitted or left editing.

## Public Payment Orders

For `对公支付单`:

- Title format should follow HuiLianYi guidance: time + project/store + payment purpose, for example `YYYY.MM供应商对公支付款`.
- Set `归属公司` to the invoice buyer and `往来公司(收款公司)` to the invoice seller.
- Prefer `到票支付` for VAT invoices and generate/import the fee from the invoice file.
- Verify seller bank account from the invoice before submission:
  - `销方开户银行`
  - `银行账号`
- If the supplier does not exist in HuiLianYi, ask before adding it.

## Bundled Resources

- `scripts/classify_and_print.py`: classify invoice PDFs/images and optionally send A5 print jobs with the standard copy rules.
- `references/huilianyi-field-guide.md`: field mapping, Excel writeback, and browser-operation details.
