// const XLSX = require('xlsx');

// // Read the file
// const workbook = XLSX.readFile('AHERI_220_14_11_2023.xlsx');

// // Get the first sheet
// const sheetName = workbook.SheetNames[0];
// const worksheet = workbook.Sheets[sheetName];

// // Convert the sheet to JSON
// const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1});

// // Get the station name from cell 'A1'
// const stationName = worksheet['A1'].v;

// // Filter out the data you need
// const data = jsonData.flatMap(row => {
//     // Identify the feeder names (with 'F' prefix)
//     const feederNames = row.filter(value => typeof value === 'string' && value.startsWith('F'));

//     return feederNames.map(feederName => ({
//         'Date': '15-10-2023', // Date is given
//         'Name of Station': stationName,
//         'Name of Feeder': feederName
//     }));
// });

// // Log the data
// console.table(data);



const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Function to process each file
function processFile(filePath) {
    // Read the file with cellDates option
    const workbook = XLSX.readFile(filePath, { cellDates: true });

    // Get the first sheet
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Convert the sheet to JSON
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Get the station name from cell 'A1'
    const stationName = worksheet['A1'].v;

    // Filter out the data you need
    const data = jsonData.flatMap((row, rowIndex) => {
        // Skip the first row (header row)
        if (rowIndex === 0) return [];

        // Identify the feeder names (with 'F' prefix)
        const feederNames = row.filter(value => typeof value === 'string' && value.startsWith('F'));

        // Get the date from the second column (column B)
        const date = row[1];

        return feederNames.map(feederName => ({
            'Date': date, // Use the date from the Excel file
            'Name of Station': stationName,
            'Name of Feeder': feederName
        }));
    });

    return data;
}
// Directory containing the Excel files
const directory = 'C:/Users/KIIT/Desktop/Code/Project/readexcel/xlsx_files'; // Adjust this path


// Array to store data from all files
const allData = [];

// Read all files in the directory
fs.readdir(directory, (err, files) => {
    if (err) {
        console.error('Error reading directory:', err);
        return;
    }

    // Filter out only Excel files
    const excelFiles = files.filter(file => path.extname(file).toLowerCase() === '.xlsx');

    // Process each Excel file sequentially
    excelFiles.forEach((file, index) => {
        const filePath = path.join(directory, file);
        const data = processFile(filePath);
        allData.push(...data);
        console.log(`Processed file: ${file}`);
    });

    // Log all data
    console.table(allData);
});



