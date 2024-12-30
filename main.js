// Required Dependencies
// Install these packages using npm:
// npm install xlsx docxtemplater pizzip fs

const fs = require('fs');
const XLSX = require('xlsx');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

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
        placeholders.push(firstRow[i]);
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
    const content = fs.readFileSync(templatePath, 'binary');
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
        console.error('Error rendering document:', error);
        return;
    }

    const buffer = doc.getZip().generate({ type: 'nodebuffer' });
    fs.writeFileSync(outputPath, buffer);
}

// Main execution
(function main() {
    const excelFile = 'data.xlsx'; // Path to Excel file
    const wordTemplate = 'template.docx'; // Path to Word template

    // Read data from the Excel file
    const clients = readExcel(excelFile);

    // Generate a document for each client
    clients.forEach((client) => {
        const outputFilePath = `${client['{{ClientName}}'] || 'Client'}.docx`; // Use the placeholder value for client name
        replacePlaceholders(wordTemplate, outputFilePath, client);
        console.log(`Generated document for client: ${outputFilePath}`);
    });
})();
