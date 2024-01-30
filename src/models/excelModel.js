// excelModel.js
const exceljs = require('exceljs');
const path = require('path');

const filePath = path.join(__dirname, '../Book4.xlsx');

const readExcelFile = async () => {
  const workbook = new exceljs.Workbook();
  await workbook.xlsx.readFile(filePath);
  return workbook.getWorksheet('User Details');
};

const writeExcelFile = async (workbook) => {
  await workbook.xlsx.writeFile(filePath);
};

module.exports = { readExcelFile, writeExcelFile };
