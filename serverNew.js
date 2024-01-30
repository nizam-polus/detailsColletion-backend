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

const filePath = path.join(__dirname, '/Book4.xlsx');

if (!fs.existsSync(filePath)) {
    const workbook = new exceljs.Workbook();
    const sheet = workbook.addWorksheet('User Details');
    sheet.addRow(['Name', 'Employee ID', 'Laptop Model', 'Mac Address', 'Registration Date', 'Last Updated Date', 'Status', "password"]);
    workbook.xlsx.writeFile(filePath)
      .then(() => {
        console.log('Excel file created successfully:', filePath);
      })
      .catch((error) => {
        console.error('Error creating Excel file:', error);
      });
  }

// Routes
const registerRoute = require('./src/routes/register');
const loginRoute = require('./src/routes/login');
const detailsRoute = require('./src/routes/details');
const updateRoute = require('./src/routes/update');
const toggleStatusRoute = require('./src/routes/toggleStatus');
const deleteRoute = require('./src/routes/delete');

app.use('/api/register', registerRoute);
app.use('/api/login', loginRoute);
app.use('/api/details', detailsRoute);
app.use('/api/update', updateRoute);
app.use('/api/toggle-status', toggleStatusRoute);
app.use('/api/delete', deleteRoute);

app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
