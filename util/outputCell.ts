import Excel from "@tinkie101/exceljs-wrapper";

async function outputSpecificCell(filename: string, sheetNumber: number, cellAddress: string) {
    const workbook = new Excel.Workbook();

    try {
        await workbook.xlsx.readFile(filename);
    } catch (error) {
        console.error("Error reading the Excel file:", error);
        return;
    }

    const worksheet = workbook.getWorksheet(sheetNumber);

    if (!worksheet) {
        console.error(`No worksheet found with index ${sheetNumber}.`);
        return;
    }

    const cell = worksheet.getCell(cellAddress);

    if (cell) {
        console.log(`Value in cell ${cellAddress}:`, cell.result);
    } else {
        console.error(`Cell ${cellAddress} not found.`);
    }
}

// Example usage
if (import.meta.main) {
    const filename = "./data/data_1.xlsx"; // Replace with your Excel file path
    const sheetNumber = 1; // Replace with the sheet number you want to access
    const cellAddress = "L4"; // Replace with the cell address you want to output

    outputSpecificCell(filename, sheetNumber, cellAddress).catch(console.error);
}