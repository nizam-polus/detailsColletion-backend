// deleteController.js
const exceljs = require('exceljs');
const path = require('path');



const deleteEntry = async (req, res) => {
  const id = req.params.id;
    const workbook = new exceljs.Workbook();
    const filePath = path.join(__dirname, '../../Book4.xlsx');

  
    workbook.xlsx.readFile(filePath)
      .then(() => {
        const sheet = workbook.getWorksheet('User Details');
        let rowToDelete = null;
  
        sheet.eachRow((row, rowNumber) => {
          if (rowNumber > 1 && row.getCell(2).value === id) {
            rowToDelete = row;
          }
        });
  
        if (rowToDelete) {
          sheet.spliceRows(rowToDelete.number, 1);
          console.log('Entry deleted in Excel file:', filePath);
          res.status(200).send('Device entry deleted successfully');
        } else {
          res.status(404).send('Device entry not found');
        }
  
        // Save the updated Excel file
        return workbook.xlsx.writeFile(filePath); 
      })
      .catch((error) => {
        console.error('Error deleting entry from Excel file:', error);
        res.status(500).send('Error deleting entry from Excel file');
      });
};

module.exports = { deleteEntry };
