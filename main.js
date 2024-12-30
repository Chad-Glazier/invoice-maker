
import { readFileSync, ensureDirSync, writeFileSync } from "https://deno.land/std@0.120.0/fs/mod.ts";
import { dirname, resolve } from "https://deno.land/std@0.120.0/path/mod.ts";
import * as XLSX from "https://deno.land/x/sheetjs/xlsx.mjs";
import PizZip from "https://deno.land/x/docxtemplater@3.26.0/mod.ts";
import Docxtemplater from "https://deno.land/x/docxtemplater@3.26.0/lib/docxtemplater.js";

// Function to read the Excel file and extract data
function readExcel(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // Convert sheet to 2D array

    const placeholders = [];
    const clients = [];

    // Extract placeholders from the first row
    const firstRow = data[0];
    for (let i = 0; i < firstRow.length; i++) {
        if (firstRow[i] === '' || firstRow[i] === undefined) break;
        placeholders.push(`((${firstRow[i]}))`); // Wrap placeholders with (( ))
    }

    // Extract client data from subsequent rows
    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === '' || data[i][0] === undefined) break;
        const clientData = {};
        for (let j = 0; j < placeholders.length; j++) {
            clientData[placeholders[j]] = data[i][j] || '';
        }
        clients.push(clientData);
    }

    return clients;
}

// Function to replace placeholders in a Word document
function replacePlaceholders(templatePath, outputPath, clientData) {
    const content = readFileSync(templatePath, { encoding: "binary" });
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });

    // Set the data for placeholders
    doc.setData(clientData);

    try {
        doc.render();
    } catch (error) {
        console.error("Error rendering document:", error);
        return;
    }

    const buffer = doc.getZip().generate({ type: "uint8array" });
    writeFileSync(outputPath, buffer);
}

// Main execution
(async function main() {
    const __dirname = dirname(new URL(import.meta.url).pathname);
    const excelFile = resolve(__dirname, "data.xlsx"); // Relative path to Excel file
    const wordTemplate = resolve(__dirname, "template.docx"); // Relative path to Word template

    // Ensure the invoices directory exists
    const invoicesDir = resolve(__dirname, "invoices");
    ensureDirSync(invoicesDir);

    // Read data from the Excel file
    const clients = readExcel(excelFile);

    // Generate a document for each client
    clients.forEach((client) => {
        const outputFilePath = resolve(invoicesDir, `${client['((Client Name))'] || 'Client'}.docx`); // Use the placeholder value for client name
        replacePlaceholders(wordTemplate, outputFilePath, client);
        console.log(`Generated document for client: ${outputFilePath}`);
    });
})();
