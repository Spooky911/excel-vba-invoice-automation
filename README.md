# Bluewater Invoice Automation

An Excel VBA project that turns orders into numbered, print-ready invoices. Customer details and product prices come from the workbook's catalogs, so an order can be built by selecting codes and entering quantities.

## Get started

1. Download [Bluewater Invoice Automation.xlsm](Bluewater%20Invoice%20Automation.xlsm) and open it in desktop Excel for Windows. Allow the workbook macros if Excel prompts you and you trust the file.
2. In **Settings**, enter your business and payment details, invoice prefix, and payment terms.
3. On **Order Sheet**, choose a customer code and product codes, then enter quantities. **Customers** and **Products** hold the source records.
4. On **Invoice**, choose **Generate Invoice**. Excel validates the order and creates an invoice with a number, issue date, due date, and calculated totals. Save the workbook to retain the invoice sequence.
5. Choose **Email Invoice** to generate an A4 PDF. A configured classic Outlook profile opens a draft with the PDF attached. Review the draft before sending.

## How it works

- Excel lookup formulas populate the customer address, billing email, product descriptions, and GBP unit prices from the selected codes.
- The VBA macro checks the populated order lines and copies a snapshot to the invoice. The generated invoice keeps its values when the order sheet changes. Generating again assigns the next invoice number.
- Line amounts are rounded to two decimal places. The discount is deducted from the VAT-inclusive subtotal. For the workbook's 20% rate, included VAT is one-sixth of the discounted total.
- The invoice uses A4 portrait, fits one page wide, repeats item headers on subsequent pages, and keeps the payment reference linked to the invoice number.

The included customer, product and email records are fictional. Enter your actual billing and payment details before issuing invoices.

## Files

| File | Purpose |
| --- | --- |
| [Bluewater Invoice Automation.xlsm](Bluewater%20Invoice%20Automation.xlsm) | Ready-to-use workbook with the catalogs, formulas, and embedded macros |
| [InvoiceAutomation.bas](InvoiceAutomation.bas) | Exported VBA module for invoice generation and email preparation |
