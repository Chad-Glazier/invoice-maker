import readExcel from "./util/readExcel1.ts"


async function generateInvoice(
	template: string,
	data: string
) {
	// gotta write this lol
	console.log(await readExcel("data/data_1.xlsx"))
}

generateInvoice("", "")
