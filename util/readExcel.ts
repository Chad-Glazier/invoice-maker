import Excel from "@tinkie101/exceljs-wrapper"
import Invoice from "./Invoice.ts";

/**
 * Used to read an Excel spreadsheet for invoice data. The following is expected:
 * - The first row of the worksheet represents the headers.
 * - The first column of each row (except the header row) contains the client's name.
 */
async function readExcel(filename: string): Promise<null |  Invoice[]> {
	// validate input, get worksheet.
	const workbook = new Excel.Workbook();
	try {
		await workbook.xlsx.readFile(filename)
	} catch {
		console.error(`${filename} does not exist.`)
		return null
	}
	const worksheet = workbook.getWorksheet(1)
	if (worksheet === undefined) {
		console.error(`${filename} exists, but no worksheet was found in it.`)
		return null
	}
	
	// this will store the output
	const invoices: Invoice[] = []

	// collect headers so that we can associate each cell's value with a key.
	const headers: string[] = []
	for (const header of worksheet.getRow(1).values as Iterable<unknown>) {
		if (header == undefined) continue
		headers.push(header.toString())
	}
	if (headers.length == 0) {
		console.error(`${filename} should have headers in its first row.`)
	}

	// iterate over the rows to build out invoice objects.
	workbook.getWorksheet(1)!.eachRow((row, rowNum) => {
		if (rowNum == 1) return // skip the header row
		const invoice = new Invoice()
		row.eachCell((cell, colNum) => {
			if (cell.value == undefined) return // skip empty cells
			invoice.set(headers[colNum - 1], cell)
		})
		invoices.push(invoice)
	});

	return invoices
}

export default readExcel
