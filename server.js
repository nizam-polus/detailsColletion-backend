const express = require('express');
const bodyParser = require('body-parser');
const exceljs = require('exceljs');
const fs = require('fs');
const path = require('path');
const cors = require('cors');
const app = express();
const PORT = process.env.PORT || 5000;

app.use(bodyParser.json());
app.use(cors()); // Enable CORS for all routes

const filePath = path.join(__dirname, '/Book3.xlsx');

if (!fs.existsSync(filePath)) {
    const workbook = new exceljs.Workbook();
    const sheet = workbook.addWorksheet('User Details');
    sheet.addRow(['Name', 'Employee ID', 'Laptop Model', 'Mac Address', 'Registration Date', 'Last Updated Date', 'Status']);
    workbook.xlsx.writeFile(filePath)
      .then(() => {
        console.log('Excel file created successfully:', filePath);
      })
      .catch((error) => {
        console.error('Error creating Excel file:', error);
      });
  }
  
  app.post('/api/register', async (req, res) => {
    const { name, employeeId, laptopModel, macAddress } = req.body;
    const currentDate = new Date().toLocaleString().replace(/,/g, '');
    const status = 'active';
  
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
        existingRow.getCell(1).value = name;
        existingRow.getCell(3).value = laptopModel;
        existingRow.getCell(4).value = macAddress;
        existingRow.getCell(6).value = currentDate;
        // existingRow.getCell(7).value = status;
        console.log('Entry updated in Excel file:', filePath);
        res.status(200).send('Device entry updated successfully');
      } else {
        // Employee ID does not exist, add a new entry
        sheet.addRow([name, employeeId, laptopModel, macAddress, currentDate, currentDate, status]);
        console.log('New registration added to Excel file:', filePath);
        res.status(200).send('New device registration added successfully');
      }
  
      // Save the updated Excel file
      await workbook.xlsx.writeFile(filePath);
    } catch (error) {
      console.error('Error updating Excel file:', error);
      res.status(500).send('Error updating Excel file');
    }
  });

  app.post('/api/login', async (req, res) => {
    const { employeeId } = req.body;
  
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
        console.log('user successfully loggined:', filePath);
        res.status(200).send('Device entry updated successfully');
      } else {
        // Employee ID does not exist, add a new entry
        
        console.log('No User Found', filePath);
        res.status(200).send('No user registerd with this employee Id');
      }
  
      // Save the updated Excel file
      await workbook.xlsx.writeFile(filePath);
    } catch (error) {
      console.error('Error updating Excel file:', error);
      res.status(500).send('Error updating Excel file');
    }
  });
  
  // New endpoint to fetch details based on Employee ID
app.get('/api/details/:id', (req, res) => {
    const id = req.params.id;
    const workbook = new exceljs.Workbook();
  
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
  });

  // update excel data
  app.put('/api/update/:id', async (req, res) => {
    const id = req.params.id;
    const updatedDetails = req.body;
  
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
    }})

// update status
  app.post('/api/toggle-status/:id', (req, res) => {
    const id = req.params.id;
    const { status } = req.body;
  
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
  });

  // Delete row
  app.delete('/api/delete/:id', (req, res) => {
    const id = req.params.id;
    const workbook = new exceljs.Workbook();
  
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
  });
  
  app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
  });