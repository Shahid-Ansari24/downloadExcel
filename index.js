const express = require('express');
const XLSX = require('xlsx');
const bodyParser = require('body-parser')
const app = express();
const port = 3000;

// Sample data
// const data = [
//   { key1: 'value1', key2: 'value2', key3: 'value3' },
//   { key1: 'value4', key2: 'value5', key3: 'value6' },
// ];

app.use(bodyParser.json())

// Convert data to Excel file
const generateExcel = (data, contentHeader) => {
    console.log("data---", data)
    console.log("contentHeader--", contentHeader)
  // Create a new workbook
  const wb = XLSX.utils.book_new();

  // Prepare the data for the worksheet
  const formattedData = [];
  if (contentHeader) {
    formattedData.push(contentHeader); // Add custom headers as the first row
  } else {
    formattedData.push(Object.keys(data[0])); // Use keys as headers if contentHeader is not provided
  }
  data.forEach(row => formattedData.push(Object.values(row))); // Add data rows

  // Convert prepared data to worksheet
  const ws = XLSX.utils.aoa_to_sheet(formattedData);

  // Append the worksheet to the workbook
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

  // Write the workbook to a binary string
  const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });

  return excelBuffer;
};

// Route to download the Excel file
app.get('/download-excel', (req, res) => {
    console.log("req.body---", req.body)
    const data = req.body.data;
    const contentHeader = req.body.ContentHeader
  // Define a contentHeader for demonstration
//   const contentHeader = ['Heading 1', 'Heading 2', 'Heading 3'];

  // Generate Excel buffer
  const excelBuffer = generateExcel(data, contentHeader);

  // Set headers for file download
  res.setHeader('Content-Disposition', 'attachment; filename=excel-file.xlsx');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(excelBuffer);
});

// Start the server
app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
