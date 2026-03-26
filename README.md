# InvoiceManager

## Project Overview

`InvoiceManager.xlsm` is an Excel macro-enabled workbook for managing invoices for a small school IT services business. It uses VBA macros to automate invoice creation, tracking, and email drafting via Microsoft Outlook.

PII (personal names, email addresses, phone numbers, and tax reference numbers) has been replaced with placeholder tokens so this repository can be used by anyone without exposing real personal or business data.

## Prerequisites

- Windows OS
- Microsoft Excel with macros enabled (the workbook is a macro-enabled `.xlsm` file)
- Microsoft Outlook installed and configured with a mail account

## Setup

1. Open `InvoiceManager.xlsm` — when prompted by Excel, choose to enable macros.
2. Go to the **Settings** sheet and configure:
   - Cell **B1** — set to your `BasePath`: the full path to the folder where school invoice folders will be created (e.g. `C:\Users\You\Invoices`)
   - Cell **B2** — set to your `CompanyName` (e.g. `Acme IT Services`)
3. Go to the **Schools** sheet and add a row for each school you work with (see [Schools Sheet column reference](#schools-sheet-column-reference) below).
4. Create the required folder structure under your `BasePath` (see [Folder Structure](#folder-structure) below).
5. After editing any `.bas` files, re-import them into `InvoiceManager.xlsm` (see [Re-importing .bas Files](#re-importing-bas-files) below).

## Schools Sheet Column Reference

| Col | Field | Example |
|-----|-------|---------|
| A | Code | TST |
| B | Name | Test School NS |
| C | Principal | Test Principal |
| D | FolderName | TestSchoolNS |
| E | Email | testprincipal@testschool.ie |
| F | SharedLink | https://example.com/shared |
| G | Phone | 01 234 5678 |
| H | Address | 1 Test Street, Dublin |
| I | CalloutFee | 75 |

## Folder Structure

The following directory layout is required under your `BasePath`:

```
BasePath\
  SchoolFolderName\
    InProgress\
    Sent\
    Paid\
    SchoolFolderName-Shared\
      Invoices\
        YYYY\
  InvoiceTemplate\
    InvoiceTemplate.xlsm
    InvoiceTemplate.xlsx
```

The macros call `EnsureFolderExists` for most paths, but the top-level school folder and the `InvoiceTemplate` folder must exist before first use.

## Usage

### CreateNewInvoice

Opens a school selection form, copies `InvoiceTemplate.xlsm` to `SchoolFolderName\InProgress\`, names it with the invoice number, and registers it in InvoiceRegister.

### EditInvoice

Opens a selection form for InProgress invoices and opens the selected invoice file for editing.

### SendInvoice

Opens a selection form for InProgress invoices, exports a PDF, moves the file to the Sent folder, and drafts an Outlook email with the PDF attached.

### MarkAsPaid

Opens a selection form for Sent invoices, prompts for a paid date, moves the file to the Paid folder, copies the PDF to the school's `Shared/Invoices/YYYY` folder, updates InvoiceRegister and TaxTracker, and drafts a payment confirmation email.

### HighlightOldSentInvoices

Highlights rows in InvoiceRegister where the invoice has been in Sent status for more than 30 days, as a visual reminder to follow up.

## Placeholder Tokens

The following tokens appear in `SendInvoice.bas` and `MarkAsPaid.bas` and must be replaced with real values before use. Until replaced, the literal placeholder text will appear in drafted emails.

- `[YOUR_NAME]` — your full name for the email signature
- `[YOUR_EMAIL]` — your email address(es) for the email signature
- `[YOUR_PHONE]` — your phone number for the email signature
- `[YOUR_TRN]` — your Tax Reference Number / PPSN (remove this line if not applicable)
- `CompanyName` (Settings sheet B2) — used in email subjects via `GetCompanyName()`

## Test School

The repo includes a `TestSchoolNS` folder with the required subfolder structure already in place.

To use it:
1. Add the test school record to the Schools sheet: Code=`TST`, Name=`Test School NS`, Principal=`Test Principal`, FolderName=`TestSchoolNS`, Email=`testprincipal@testschool.ie`, SharedLink=`https://example.com/shared`, Phone=`01 234 5678`, Address=`1 Test Street Dublin`, CalloutFee=`75`
2. Set `BasePath` in Settings!B1 to the repo root (or wherever `TestSchoolNS` lives)
3. Run `CreateNewInvoice` and select `TST` to verify the system works end-to-end

Once you have added your own school data, you can delete the `TestSchoolNS` folder and remove the `TST` row from the Schools sheet.

## Re-importing .bas Files

After editing any `.bas` file:

1. Open `InvoiceManager.xlsm`
2. Press `Alt+F11` to open the VBA IDE
3. In the Project Explorer, right-click the module you want to replace and select "Remove [ModuleName]" (choose "No" when asked to export)
4. Go to File → Import File and select the edited `.bas` file
5. Save and close the VBA IDE
6. Save `InvoiceManager.xlsm`
