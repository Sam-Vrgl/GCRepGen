const XLSX = require('xlsx');

// Create a new workbook
const wb = XLSX.utils.book_new();

// Sample data
const data = [
    { Name: "Alice", Age: 30, City: "New York" },
    { Name: "Bob", Age: 25, City: "Los Angeles" },
    { Name: "Charlie", Age: 35, City: "Chicago" }
];

// Create a worksheet from the data
const ws = XLSX.utils.json_to_sheet(data);

// Add the worksheet to the workbook
XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

// Write the workbook to a file
XLSX.writeFile(wb, "sample.xlsx");

console.log("sample.xlsx created successfully!");
