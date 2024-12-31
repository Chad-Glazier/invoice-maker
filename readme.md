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
