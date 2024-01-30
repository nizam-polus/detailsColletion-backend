// detailsController.js
const exceljs = require('exceljs');
const path = require('path');


const getDetails = async (req, res) => {
  const id = req.params.id;
    const workbook = new exceljs.Workbook();
    const filePath = path.join(__dirname, '../../Book4.xlsx');

  
    workbook.xlsx.readFile(filePath)
      .then(() => {
        const sheet = workbook.getWorksheet('User Details');
  
        let details;
        sheet.eachRow((row, rowNumber) => {
          if (rowNumber > 1 && row.getCell(2).value === id) {
            details = {
              name: row.getCell(1).value,
              employeeId: row.getCell(2).value,
              laptopModel: row.getCell(3).value,
              macAddress: row.getCell(4).value,
              registrationDate: row.getCell(5).value,
              lastUpdatedDate: row.getCell(6).value,
              status: row.getCell(7).value,
              password: row.getCell(8).value
            };
          }
        });
  
        if (details) {
          res.status(200).send(details);
        } else {
          res.status(404).send('Details not found');
        }
      })
      .catch((error) => {
        console.error('Error fetching details:', error);
        res.status(500).send('Error fetching details');
      });
};

module.exports = { getDetails };
