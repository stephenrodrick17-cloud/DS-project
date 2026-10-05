const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const filePath = path.join(__dirname, 'products_database.xlsx');

const sampleProducts = [
  { ID: 1, Name: 'Wireless Mouse', Category: 'Electronics', 'Price (INR)': 799, Stock: 45, Rating: 4.3, Status: 'In Stock' },
  { ID: 2, Name: 'Mechanical Keyboard', Category: 'Electronics', 'Price (INR)': 2999, Stock: 18, Rating: 4.6, Status: 'In Stock' },
  { ID: 3, Name: 'USB-C Charger 65W', Category: 'Electronics', 'Price (INR)': 1499, Stock: 60, Rating: 4.4, Status: 'In Stock' },
  { ID: 4, Name: 'Bluetooth Speaker', Category: 'Electronics', 'Price (INR)': 2199, Stock: 0, Rating: 4.2, Status: 'Out of Stock' },
  { ID: 5, Name: 'Cotton T-Shirt', Category: 'Clothing', 'Price (INR)': 499, Stock: 120, Rating: 4.1, Status: 'In Stock' },
  { ID: 6, Name: 'Denim Jacket', Category: 'Clothing', 'Price (INR)': 2499, Stock: 12, Rating: 4.5, Status: 'Low Stock' },
  { ID: 7, Name: 'Running Shoes', Category: 'Footwear', 'Price (INR)': 3299, Stock: 30, Rating: 4.7, Status: 'In Stock' },
  { ID: 8, Name: 'Leather Wallet', Category: 'Accessories', 'Price (INR)': 899, Stock: 8, Rating: 4.0, Status: 'Low Stock' },
  { ID: 9, Name: 'Steel Water Bottle', Category: 'Home', 'Price (INR)': 649, Stock: 75, Rating: 4.5, Status: 'In Stock' },
  { ID: 10, Name: 'Desk Lamp', Category: 'Home', 'Price (INR)': 1199, Stock: 0, Rating: 4.3, Status: 'Out of Stock' },
  { ID: 11, Name: 'Notebook (Pack of 5)', Category: 'Stationery', 'Price (INR)': 350, Stock: 200, Rating: 4.2, Status: 'In Stock' },
  { ID: 12, Name: 'Gel Pen Set', Category: 'Stationery', 'Price (INR)': 180, Stock: 150, Rating: 4.4, Status: 'In Stock' }
];

const readmeData = [
  { 'How to use this file': "Sheet 'Products' is the database. Row 1 = column headers (do not rename them)." },
  { 'How to use this file': 'Edit any value in rows 2 onward - the website will show the change after reload/refresh.' },
  { 'How to use this file': 'Add a new product by adding a new row at the bottom; delete a row to remove one.' },
  { 'How to use this file': 'Keep ID unique. Status should be: In Stock, Low Stock or Out of Stock.' },
  { 'How to use this file': 'Row 2 (Wireless Mouse) is an example of the expected format.' },
  { 'How to use this file': 'Sample data is made up for practice.' }
];

function createDatabase() {
  const wb = xlsx.utils.book_new();
  const wsProducts = xlsx.utils.json_to_sheet(sampleProducts);
  const wsReadme = xlsx.utils.json_to_sheet(readmeData);

  xlsx.utils.book_append_sheet(wb, wsProducts, 'Products');
  xlsx.utils.book_append_sheet(wb, wsReadme, 'README');

  xlsx.writeFile(wb, filePath);
  console.log(`Successfully generated/reset ${filePath}`);
}

createDatabase();
