// toggleStatusController.js
const exceljs = require('exceljs');
const path = require('path');


const toggleStatus = async (req, res) => {
  const id = req.params.id;
  const { status } = req.body;
  const filePath = path.join(__dirname, '../../Book4.xlsx');

  try {
    const workbook = new exceljs.Workbook();
    workbook.xlsx.readFile(filePath)
      .then(() => {
        const sheet = workbook.getWorksheet('User Details');

        let existingRow;
        sheet.eachRow((row, rowNumber) => {
          if (rowNumber > 1 && row.getCell(2).value === id) {
            existingRow = row;
          }
        });

        if (existingRow) {
          existingRow.getCell(7).value = status;
          return workbook.xlsx.writeFile(filePath);
        } else {
          res.status(404).send('User not found');
        }
      })
      .then(() => {
        console.log(`User status toggled to ${status} in Excel file:`, filePath);
        res.status(200).send(`User status toggled to ${status} successfully`);
      })
      .catch((error) => {
        console.error('Error toggling user status:', error);
        res.status(500).send('Error toggling user status');
      });
  } catch (error) {
    console.error('Error toggling user status:', error);
    res.status(500).send('Error toggling user status');
  }
};

module.exports = { toggleStatus };
