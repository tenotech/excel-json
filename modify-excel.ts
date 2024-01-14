import * as ExcelJS from 'exceljs';

async function updateCell(filePath: string, sheetName: string, rowIndex: number, columnIndex: number, newValue: any) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) {
    console.error(`Sheet '${sheetName}' not found in the Excel file.`);
    return;
  }

  const cell = sheet.getCell(rowIndex, columnIndex);
  if (!cell) {
    console.error(`Cell at row ${rowIndex}, column ${columnIndex} not found in the sheet.`);
    return;
  }

  cell.value = newValue;

  await workbook.xlsx.writeFile(filePath);
  console.log(`Cell at row ${rowIndex}, column ${columnIndex} updated successfully.`);
}

// Example usage
const filePath = './myExcel.xlsx';
const sheetName = 'Marketplace_Marchands';  // Replace with your sheet name
const rowIndex = 2;           // Replace with the row index (1-based)
const columnIndex = 2;        // Replace with the column index (1-based)
const newValue = 'New Value'; // Replace with the new value

updateCell(filePath, sheetName, rowIndex, columnIndex, newValue);
