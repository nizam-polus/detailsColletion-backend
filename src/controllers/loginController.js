// loginController.js
const exceljs = require('exceljs');
const path = require('path');


const login = async (req, res) => {
  const { employeeId, password } = req.body;
  const filePath = path.join(__dirname, '../../Book4.xlsx');


  try {
    const workbook = new exceljs.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet('User Details');

    let existingRow = null;

    // Check if there is a row with the specified Employee ID and password
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1 && row.getCell(2).value === employeeId && row.getCell(8).value === password) {
        existingRow = row;
      }
    });

    if (existingRow) {
      console.log('User successfully logged in:', filePath);
      res.status(200).send('User logged in successfully');
    } else {
      console.log('No User Found', filePath);
      res.status(401).send('Invalid credentials');
    }
  } catch (error) {
    console.error('Error updating Excel file:', error);
    res.status(500).send('Error updating Excel file');
  }
};

module.exports = { login };
