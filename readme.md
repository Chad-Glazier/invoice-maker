To run the code, run

```ps
deno run main.ts
```

in PowerShell.

## File Structure

`data` has the Excel files that contain the data used to populate a given
template.

`templates` has the `.docx` (MS Word) files that represent invoice templates.

`generated_invoices` contains the output invoices.

`util` is a group of utility functions used in the main program.

`deno.json` and `deno.lock` contain information for the deno runtime to manage
external dependencies and scripts.

## Dependencies

The [docxml](https://deno.land/x/docxml@5.15.3) package is used to interact with `.docx` (MS Word) files.

```ts
import * as docxml from "https://deno.land/x/docxml@5.15.3/mod.ts";
```

The [exceljs](https://jsr.io/@tinkie101/exceljs-wrapper) package is used to interact with `.xlsx` (MS Excel) files.

```ts
import * as exceljs_wrapper from "@tinkie101/exceljs-wrapper"
```
