# HuiLianYi Field Guide

## Invoice Extraction

Extract these fields from each PDF:

- Invoice kind: `电子发票（增值税专用发票）`, `电子发票（普通发票）`, foreign invoice, or payment evidence.
- Buyer: `购买方` / `购方` name and taxpayer id.
- Seller: `销售方` / `销方` name and taxpayer id.
- Amount: `价税合计（小写）`.
- Seller bank: `销方开户银行` and `银行账号`.
- Date and invoice number.

If `pypdf` text extraction is poor, use rendered screenshots or open the PDF visually before classifying.

## Excel Workbook Pattern

Preferred sheets:

- `汇总`: grouped by `发票抬头`, with report id/type/amount and notes.
- `按抬头明细`: row-level source of truth.
- One sheet per invoice title when helpful.
- `无票`: no-ticket expenses and payment proofs.
- `报销单回写`: concise report-id crosswalk.
- `提交清单`: attachment-ready table in the user's required format.
- `报销一览表`: audit-facing overview that combines invoice/no-ticket evidence, fixed-asset ownership, HuiLianYi report ids, report totals, and remarks.

Submission list columns:

1. `发票抬头`
2. `费用名称`
3. `资产编号`
4. `使用人`
5. `使用部门`
6. `金额`
7. `采购原因`

Always verify that the `采购原因` column is not accidentally shifted into a report-number column after adding writeback columns.

Overview table columns:

1. `发票抬头`
2. `费用名称`
3. `资产编号`
4. `使用人`
5. `使用部门`
6. `金额`
7. `汇联易报销单号`
8. `采购原因`
9. `报销单类型`
10. `汇联易单据金额`
11. `报销备注`

Use `凭证费用` or the actual payer/counterparty as the visible title for payment screenshots or other no-ticket evidence. Avoid putting `无票` in editable workbook fields that will be uploaded or copied into HuiLianYi.
For payment-proof items, set the internal `报销单类型` to the HuiLianYi no-ticket document type when needed for reconciliation; for invoice-backed personal reimbursement, use `日常报销单`; for public payment, use `对公支付单`.

## Reimbursement Rules Captured From User Preference

- Before browser entry, classify each item as `个人报销` or `对公付款`.
- Personal reimbursement pays the logged-in applicant; do not route personal reimbursement money to suppliers.
- Public payment pays the invoice issuer or contract counterparty; do not route public-payment money to the logged-in applicant.
- HuiLianYi titles must use Chinese Han characters rather than pinyin. If a brand or vendor is written in Latin letters, use a Chinese purpose phrase or official Chinese company short name in the title.
- Same invoice title can be submitted under one expense/report when reasonable.
- Payment-proof evidence without a VAT invoice must be entered through HuiLianYi's no-ticket/hand-entered fee flow, not as invoice reimbursement.
- Do not type `无票` into editable titles,事由说明, remarks, uploaded file names, or user-facing list names. Use `凭证费用`, `支付凭证费用`, `票据缺失类费用`, `支付凭证`, or `小票` according to the screen.
- When the user asks to rebuild, first compare the new source totals with prior validated outputs and existing HuiLianYi report totals. If the old total is larger, stop and identify missing rows instead of creating a smaller replacement report.
- Upload files directly rather than importing zip packages.
- Attach an updated submission list to the relevant reimbursement when the list content changes.
- Separate reports by title when that improves audit clarity; combine same-title expenses when the user allows.

## Document Type Decision

Use this decision before creating a HuiLianYi document:

| Evidence / intent | HuiLianYi document | Payee | Title pattern |
| --- | --- | --- | --- |
| Employee already paid and requests reimbursement | `个人报销` / daily reimbursement | Logged-in applicant | `YYYY年MM月<中文事项>费用报销` |
| Company should pay the invoice seller | `对公支付单` | Invoice issuer | `YYYY年MM月<中文收款方或事项>对公付款` |
| Company should pay a contract counterparty | `对公支付单` | Contract counterparty | `YYYY年MM月<中文合同方或事项>对公付款` |
| Payment screenshot / receipt without VAT invoice | no-ticket hand-entered reimbursement | Logged-in applicant for personal reimbursement | `YYYY年MM月凭证费用报销` |
| Payer or recipient is ambiguous | Stop and ask | Unknown | Do not create the document yet |

## Rebuild And Total Reconciliation

Before browser entry in a rebuild or video-recording run:

1. Inventory every source the user has already provided: current workbook, previous classified workbook, previous submission list, existing HuiLianYi report list, Desktop invoice folder, WeWork cache PDFs/images, and newly uploaded screenshots.
2. Normalize all evidence rows into `发票抬头/费用名称/金额/采购原因/报销单类型/证据文件`.
3. Compare totals by report type, visible title, and fee purpose. Use `scripts/reconcile_reimbursement_totals.py` when both sides are in CSV/XLSX form.
4. If an existing HuiLianYi report or prior validated workbook has a larger total than the rebuilt source, list the missing items and pause. Do not create or save a smaller replacement report unless the user confirms the reduction.
5. If `重新建一套` is requested, create new reports and keep old draft/report numbers only as comparison evidence. Reusing an existing zero-amount draft requires explicit user approval.

Observed failure modes to prevent:

- A rebuilt payment-proof report can drop screenshots that were not in the newest workbook but were present in WeWork cache attachments or an older submission list.
- HuiLianYi may reject editable content containing `无票`; the fixed system type can remain, but user-entered text must use neutral wording.
- Creating from an old zero-amount draft can look like a new report but silently inherit stale state.

## Fixed-Asset Matching And Overview Table

When invoices or no-ticket screenshots need to be matched to fixed-asset rows:

1. Extract evidence fields from every invoice or screenshot: invoice title or payment-proof marker, fee name, amount, payment purpose, supplier or counterparty, and evidence file name.
2. Match each evidence row to the fixed-asset workbook using this priority:
   - exact amount match within 0.01 plus a close fee-name match;
   - exact asset number when the evidence or workbook already contains it;
   - close purchase-reason or department match when multiple rows have the same amount;
   - manual review when there are duplicate amounts and weak text matches.
3. Fill `资产编号`, `使用人`, and `使用部门` from the matched fixed-asset row. If no match is safe, leave the asset fields blank and put `资产匹配待核对` in `报销备注`.
4. After HuiLianYi ids are known, write `汇联易报销单号`, `报销单类型`, and `汇联易单据金额` back into the overview table.
5. Merge repeated `汇联易报销单号` and `汇联易单据金额` cells only for contiguous rows that belong to the same report. Keep line-level `金额` unmerged.

Use `scripts/build_reimbursement_overview.py` when the source evidence/asset rows have already been extracted to CSV or XLSX. The script accepts a normalized evidence table and an optional fixed-asset table, then outputs the audit-facing `报销一览表` format.

For visible workbook/upload column names, prefer `凭证费用` over `无票` for payment-proof rows when the table may be uploaded to HuiLianYi. Keep any system-only no-ticket wording in internal notes or reconciliation fields.

## HuiLianYi No-Ticket / Payment-Proof Entry

For evidence that is only a screenshot, receipt, or foreign payment proof:

1. Create a hand-entered/no-ticket fee line with the correct expense type.
2. If HuiLianYi requires an invoice-like record, choose `手录发票` / `小票` rather than VAT invoice upload. Fill the total, tax-inclusive total, date, and non-deductible or zero-tax values according to the form.
3. Allocate the full manual receipt amount to the related fee rows.
4. Upload payment screenshots, receipts, and the current submission list as attachments.
5. Save, wait for asynchronous save completion, refresh if needed, and verify the missing-invoice warning or empty invoice field is cleared.
6. Keep editable text neutral: `凭证费用`, `支付凭证`, `小票`, or the actual fee purpose. Do not use `无票` in typed content.

## HuiLianYi Expense Types Observed

- AI/subscriptions: `办公软件/会员`.
- Payment screenshot without VAT invoice: no-ticket hand-entered fee.
- Fixed asset-like electronics: `电子设备(3年)` may require `固定资产编码`.
- Low-value equipment where no fixed asset code is available: `低值易耗电子/器具/设备`.
- Phone charges: `电信费`.
- Certificates/corporate verification: `咨询服务/证照`.

## Public Payment Example

For a vendor invoice like:

- Buyer: `<invoice buyer company>`
- Seller: `<invoice seller company>`
- Amount: `<invoice total amount>`
- Seller bank: `<seller bank name>`
- Seller account: `<seller bank account>`

Create `对公支付单`, set buyer as `归属公司`, seller as `往来公司(收款公司)`, use `到票支付`, generate the fee from invoice, and pause if `电子设备(3年)` asks for a fixed asset code.

## Final Verification

Before final response:

- Browser amount equals Excel amount.
- Report line count matches expected items.
- Newly generated report/expense numbers are written back to Excel when requested.
- Report status is exactly what user asked for: `编辑中`, `审批中`, submitted, or paid.
- Printing jobs show queue ids when printing was requested.
