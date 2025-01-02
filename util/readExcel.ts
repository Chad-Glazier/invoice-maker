import Excel from "@tinkie101/exceljs-wrapper"

async function readExcel(filename: string) {
	const workbook = new Excel.Workbook();

	try {
		await workbook.xlsx.readFile(filename);
	} catch {
		console.error("poopity scoop (the excel file does not exist)")
	}
	
	const map: Map<string, Map<string, string>> = new Map()

	const headers: string[] = []
	for (let header of workbook.getWorksheet(1)!.getRow(1).values as Iterable<any>) {
		if (header == undefined) continue
		headers.push(header)
	}

	console.log(headers)

	const clients: string[] = []
	for (let client of workbook.getWorksheet(1)!.getColumn(1).values.slice(2) as Iterable<any>) {
		if (client == undefined) continue
		clients.push(client)
	}

	console.log(clients)

	// not done
}

export default readExcel