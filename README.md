# IKO_LOADER-AUTOMATION

This repository contains a VBA macro for exporting multiple worksheets as a single PDF invoice. The code resides in `macros/IKO_CREATE_PDF_Invoices.vba`.

## Macro Overview

The macro looks for sheet names listed in column `AN` of the `Information` sheet. For each row:

1. It splits the comma-separated list of sheet names.
2. It verifies that every sheet exists in the workbook.
3. The invoice number is taken from cell `G5` on the first sheet in the list.
4. The referenced sheets are selected and exported together as a PDF using that invoice number as the filename.

Invalid groups are skipped with a warning message. After processing all rows, a completion message appears.
