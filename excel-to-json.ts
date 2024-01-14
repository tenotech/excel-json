import * as ExcelJS from 'exceljs';
import * as fs from 'fs';

interface DataPath {
  excel: string;
  json: string;
}

const dataPath:DataPath = JSON.parse(fs.readFileSync('paths.json', 'utf-8'));

async function excelToJson(filePath: string): Promise<{ [sheetName: string]: Array<{ coordinates: { row: number, column: number }, data: any }> }> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const result: { [sheetName: string]: Array<{ coordinates: { row: number, column: number }, data: any }> } = {};

  workbook.eachSheet(sheet => {
    const sheetData: Array<{ coordinates: { row: number, column: number }, data: any }> = [];

    sheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        sheetData.push({
          coordinates: {
            row: rowNumber,
            column: colNumber,
          },
          data: cell.value,
        });
      });
    });

    result[sheet.name] = sheetData;
  });

  return result;
}

// Example usage
const filePath = dataPath.excel;
const jsonPath = dataPath.json;

excelToJson(filePath)
  .then(data => {
    const jsonOutput = JSON.stringify(data, null, 2);
    fs.writeFileSync(jsonPath, jsonOutput);
    console.log('Data extracted and saved to output.json');
  })
  .catch(error => console.error('Error:', error));
