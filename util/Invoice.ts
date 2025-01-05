import Excel from "@tinkie101/exceljs-wrapper";
import { GST } from "./globals.ts"

function similar(str1: string, ...matchStrings: string[]): boolean {
	for (const matchString of matchStrings) {
		if (str1.toLowerCase().includes(matchString.toLowerCase()))
			return true
	}
	return false
}

function cellError(cell: Excel.Cell, message: string): Error {
	return new Error(`Error in row ${cell.row}, column ${cell.col}: ${message}`)
}

function requireNumber(header: string, cell: Excel.Cell, emptyAllowed: boolean = true): number | undefined {
	if (cell.formula != undefined) {
		throw cellError(cell, `formulae not allowed for ${header} values.`)
	}
	try {
		return parseFloat(cell.text)
	} catch {
		if (emptyAllowed && cell.text == "") {
			return undefined
		}
		throw cellError(cell, `numeric value expected for ${header} values.`)
	}
}

function requireString(header: string, cell: Excel.Cell, emptyAllowed: boolean = true): string | undefined {
	if (cell.formula != undefined) {
		throw cellError(cell, `formulae not allowed for ${header} values.`)
	}
	if (cell.text == "" && !emptyAllowed) {
		throw cellError(cell, `values must be provided for ${header} values.`)
	}
	if (cell.text == "") {
		return undefined
	}
	return cell.text
}

class Invoice {
	client: string
	address?: string
	city?: string
	invoiceNumber?: number
	postalCode?: string
	items: {
		description: string
		price: number
	}[]

	constructor() {
		this.client = ""
		this.address = ""
		this.city = ""
		this.invoiceNumber = 0
		this.items = []
	}

	get subtotal(): number {
		let subtotal = 0
		this.items.forEach(({ price }) => subtotal += price)
		return subtotal
	}

	get gst(): number {
		return this.subtotal * GST
	}

	get total(): number {
		return this.subtotal + this.gst
	}

	/**
	 * Will attempt to match the header to one of the known properties of an
	 * invoice. If no matching property is found, the unknown header will be
	 * logged to the console.
	 * 
	 * @param header the string representing the header for the cell's column.
	 * @param cell the cell holding the associated value.
	 */
	set(header: string, cell: Excel.Cell) {
		if (similar(header, "client")) {
			this.client = requireString(header, cell, false)!
			return
		}
		if (similar(header, "address")) {
			this.address = requireString(header, cell)
			return
		}
		if (similar(header, "city", "town")) {
			this.city = requireString(header, cell)
			return
		}
		if (similar(header, "invoice", "number", "id")) {
			this.invoiceNumber = requireNumber(header, cell)
			return
		}
		if (similar(header, "postal", "zip")) {
			this.postalCode = requireString(header, cell)
			return
		}
		if (similar(header, "description", "item")) {
			const description = requireString(header, cell)
			if (description == undefined) {
				return
			}
			this.items.push({
				description,
				price: NaN
			})
			return
		}
		if (similar(header, "price")) {
			const price = requireNumber(header, cell, true)
			if (price == undefined) return
			if (this.items.length == 0) {
				throw cellError(cell, `price provided before any item description.`)
			}
			if (!isNaN(this.items[this.items.length - 1].price)) {
				throw cellError(cell, `price provided before a new item description.`)
			}
			this.items[this.items.length - 1].price = price
			return
		}
		if (similar(header, "total", "subtotal", "gst")) {
			return
		}
		console.log(`Ignoring unknown header ${header}`)
	}
}

export default Invoice