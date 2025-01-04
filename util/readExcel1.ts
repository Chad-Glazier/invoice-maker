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

    const getCellValue = (cell: any): string => {
        if (cell.formula) {
            
            return cell.result ? cell.result.toString() : "";
        }
        return cell.value ? cell.value.toString() : "";
    };

    worksheet.eachRow((row, rowIndex) => {
        if (rowIndex === 1) return;

        const clientNameCell = row.getCell(1);
        const clientName = getCellValue(clientNameCell);
        if (!clientName) return;

        const clientData = new Map<string, string>();
        headers.forEach((header, colIndex) => {
            const cell = row.getCell(colIndex + 1);
            clientData.set(header, getCellValue(cell));
        });

        map.set(clientName, clientData);
    });

    return map;
}

export default readExcel1;
