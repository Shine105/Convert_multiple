const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// Function to convert an XLS file to XLSX
function convertXLSX(inputFilePath, outputFilePath) {
    const workbook = XLSX.readFile(inputFilePath, { cellDates: true });
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, workbook.Sheets[workbook.SheetNames[0]], 'Sheet1');

    // Create the directory if it doesn't exist
    const outputDir = path.dirname(outputFilePath);
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
    }

    XLSX.writeFile(newWorkbook, outputFilePath);
}

// Directory containing the XLS files
const directory = 'C:/Users/KIIT/Desktop/Code/Project/readexcel/BGK_14112023'; // Adjust this path

// Read all files in the directory
fs.readdir(directory, (err, files) => {
    if (err) {
        console.error('Error reading directory:', err);
        return;
    }

    // Filter out only XLS files
    const xlsFiles = files.filter(file => path.extname(file).toLowerCase() === '.xls');

    // Convert each XLS file to XLSX
    xlsFiles.forEach((file, index) => {
        const inputFilePath = path.join(directory, file);
        const outputFilePath = `./xlsx_files/output_${index}.xlsx`; // Output file name
        convertXLSX(inputFilePath, outputFilePath);
        console.log(`Converted ${file} to XLSX: ${outputFilePath}`);
    });
});
