import * as ExcelJS from 'exceljs';
import * as fs from 'fs';

interface DataPath {
    excel: string;
    json: string;
  }

const dataPath:DataPath = JSON.parse(fs.readFileSync('paths.json', 'utf-8'));

async function updateCellFromJson(filePath: string, jsonPath: string) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const jsonData = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));

  for (const sheetName in jsonData) {
    if (jsonData.hasOwnProperty(sheetName)) {
      const sheetData = jsonData[sheetName];

      const sheet = workbook.getWorksheet(sheetName);
      if (!sheet) {
        console.error(`Sheet '${sheetName}' not found in the Excel file.`);
        continue;
      }

      for (const item of sheetData) {
        const { coordinates, data } = item;
        const { row, column } = coordinates;

        const cell = sheet.getCell(row, column);
        if (!cell) {
          console.error(`Cell at row ${row}, column ${column} not found in the sheet.`);
          continue;
        }

        cell.value = data;
      }
    }
  }

  await workbook.xlsx.writeFile(filePath);
  console.log('Data updated successfully.');
}

// Example usage
const excelFilePath = dataPath.excel;
const jsonFilePath = dataPath.json;

updateCellFromJson(excelFilePath, jsonFilePath);
