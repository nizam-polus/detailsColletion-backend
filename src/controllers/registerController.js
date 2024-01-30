// registerController.js
const exceljs = require('exceljs');
const path = require('path');


const register = async (req, res) => {
  const { name, employeeId, laptopModel, macAddress, password } = req.body;
  const currentDate = new Date().toLocaleString().replace(/,/g, '');
  const status = 'active';
  const filePath = path.join(__dirname, '../../Book4.xlsx');

  try {
    const workbook = new exceljs.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet('User Details');

    let existingRow = null;

    // Check if there is a row with the specified Employee ID
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1 && row.getCell(2).value === employeeId) {
        existingRow = row;
      }
    });

    if (existingRow) {
      // Employee ID exists, update the existing row

      // existingRow.getCell(1).value = name;
      // existingRow.getCell(3).value = laptopModel;
      // existingRow.getCell(4).value = macAddress;
      // existingRow.getCell(6).value = currentDate;
      // existingRow.getCell(7).value = status;
      // existingRow.getCell(8).value = password; // Move password to the last column
      // console.log('Entry updated in Excel file:', filePath);
      // res.status(200).send('Device entry updated successfully');
      console.log("alredy registerd")
      res.status(200).send('Already Registerd With This employee id')
    } else {
      // Employee ID does not exist, add a new entry
      sheet.addRow([name, employeeId, laptopModel, macAddress, currentDate, currentDate, status, password]); // Move password to the last column
      console.log('New registration added to Excel file:', filePath);
      res.status(200).send('New device registration added successfully');
    }

    // Save the updated Excel file
    await workbook.xlsx.writeFile(filePath);
  } catch (error) {
    console.error('Error updating Excel file:', error);
    res.status(500).send('Error updating Excel file');
  }
};

module.exports = { register };
