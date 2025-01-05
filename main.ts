import readExcel1 from "./util/readExcel1.ts"
import readExcel from "./util/readExcel.ts"

async function generateInvoice(
	template: string,
	data: string
) {
	const invoices = await readExcel(data)

	console.log(invoices)
}

generateInvoice("", "data/data_1.xlsx")
