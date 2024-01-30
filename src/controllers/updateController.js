// updateController.js
const exceljs = require('exceljs');
const path = require('path');


const updateDetails = async (req, res) => {
  const id = req.params.id;
  const updatedDetails = req.body;
  const filePath = path.join(__dirname, '../../Book4.xlsx');

  try {
    const workbook = new exceljs.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet('User Details');

    let existingRow = null;

    // Check if there is a row with the specified Employee ID
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1 && row.getCell(2).value === id) {
        existingRow = row;
      }
    });

    if (existingRow) {
      // Update the existing row with the edited details
      existingRow.getCell(1).value = updatedDetails.name;
      existingRow.getCell(3).value = updatedDetails.laptopModel;
      existingRow.getCell(4).value = updatedDetails.macAddress;
      existingRow.getCell(6).value = new Date().toLocaleString().replace(/,/g, '');
      existingRow.getCell(7).value = 'active';
      existingRow.getCell(8).value = updatedDetails.password

      console.log('Entry updated in Excel file:', filePath);
      res.status(200).send('Device entry updated successfully');
    } else {
      res.status(404).send('Entry not found');
    }

    // Save the updated Excel file
    await workbook.xlsx.writeFile(filePath);
  } catch (error) {
    console.error('Error updating Excel file:', error);
    res.status(500).send('Error updating Excel file');
  }
};

module.exports = { updateDetails };
