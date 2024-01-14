import * as XLSX from 'xlsx';
import * as fs from 'fs';

const excelFilePath = './myExcel.xlsx';
const jsonFilePath = 'output.json';

// Read the existing Excel file to preserve formatting
const existingWorkbook = XLSX.readFile(excelFilePath);
const existingSheet = existingWorkbook.Sheets[existingWorkbook.SheetNames[0]];

// Read the modified JSON data
const jsonData = JSON.parse(fs.readFileSync(jsonFilePath, 'utf-8'));

// Create a new worksheet with the existing formatting
const ws = XLSX.utils.aoa_to_sheet(jsonData);
XLSX.utils.book_append_sheet(existingWorkbook, ws, 'Sheet1');

// Write the updated workbook back to the Excel file
XLSX.writeFile(existingWorkbook, excelFilePath);

console.log('Excel file updated from JSON successfully.');
