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

Submission list columns:

1. `发票抬头`
2. `费用名称`
3. `资产编号`
4. `使用人`
5. `使用部门`
6. `金额`
7. `采购原因`

Always verify that the `采购原因` column is not accidentally shifted into a report-number column after adding writeback columns.

## Reimbursement Rules Captured From User Preference

- Before browser entry, classify each item as `个人报销` or `对公付款`.
- Personal reimbursement pays the logged-in applicant; do not route personal reimbursement money to suppliers.
- Public payment pays the invoice issuer or contract counterparty; do not route public-payment money to the logged-in applicant.
- HuiLianYi titles must use Chinese Han characters rather than pinyin. If a brand or vendor is written in Latin letters, use a Chinese purpose phrase or official Chinese company short name in the title.
- Same invoice title can be submitted under one expense/report when reasonable.
- No-ticket evidence must be entered as `无票报销`, not as invoice reimbursement.
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
| Payer or recipient is ambiguous | Stop and ask | Unknown | Do not create the document yet |

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
