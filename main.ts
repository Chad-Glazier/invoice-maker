import readExcel1 from "./util/readExcel1.ts"


async function generateInvoice(
	template: string,
	data: string
) {
	// gotta write this lol
	console.log(await readExcel1("data/data_1.xlsx"))
}

generateInvoice("", "")
