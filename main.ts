import readExcel1 from "./util/readExcel1.ts"
import readExcel from "./util/readExcel.ts"

async function generateInvoice(
	template: string,
	data: string
) {
	// gotta write this lol
	console.log(await readExcel("data/data_1.xlsx"))
}

generateInvoice("", "")
