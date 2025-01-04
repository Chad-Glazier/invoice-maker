import Excel from "@tinkie101/exceljs-wrapper";

async function readExcel1(filename: string) {
    const workbook = new Excel.Workbook();

    try {
        await workbook.xlsx.readFile(filename);
    } catch {
        console.error("poopity scoop (the excel file does not exist)");
        return;
    }

    const map: Map<string, Map<string, string>> = new Map();

    
    const headers: string[] = [];
    const headerRow = workbook.getWorksheet(1)!.getRow(1).values as Array<string>;
    for (const header of headerRow) {
        if (header) {
            headers.push(header.toString());
        }
    }

    console.log("Headers:", headers);

    
    const worksheet = workbook.getWorksheet(1);
    worksheet.eachRow((row, rowIndex) => {
        if (rowIndex === 1) return;

        const clientName = row.getCell(1).value;
        if (!clientName) return;

        const clientData = new Map<string, string>();
        headers.forEach((header, colIndex) => {
            const cellValue = row.getCell(colIndex + 1).value;
            clientData.set(header, cellValue ? cellValue.toString() : "");
        });

        map.set(clientName.toString(), clientData);
    });

    return map;
}

export default readExcel1;

// returns map with [object object] where formulas are in the excel file