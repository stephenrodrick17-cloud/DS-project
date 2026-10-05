const express = require('express');
const path = require('path');
const fs = require('fs');
const cors = require('cors');
const xlsx = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;
const EXCEL_FILE_PATH = path.join(__dirname, 'products_database.xlsx');
const SHEET_NAME = 'Products';

// Enable CORS and JSON parsing
app.use(cors());
app.use(express.json());

// Serve static frontend files from 'public' folder
app.use(express.static(path.join(__dirname, 'public')));

// Cache for graceful fallback when Excel holds a transient file lock during save
let cachedData = null;
let lastModifiedTime = null;

/**
 * Helper to read and parse Excel file with retry logic for transient Windows locks
 */
function readExcelFile(maxRetries = 2, delayMs = 120) {
  let attempts = 0;

  while (attempts <= maxRetries) {
    try {
      if (!fs.existsSync(EXCEL_FILE_PATH)) {
        const error = new Error(`Database file not found: ${path.basename(EXCEL_FILE_PATH)}`);
        error.code = 'ENOENT';
        throw error;
      }

      const stat = fs.statSync(EXCEL_FILE_PATH);

      // Read Excel workbook from disk on every request
      // Using fs.readFileSync ensures fresh buffer from disk
      const fileBuffer = fs.readFileSync(EXCEL_FILE_PATH);
      const workbook = xlsx.read(fileBuffer, { type: 'buffer', cellDates: true });

      if (!workbook.Sheets[SHEET_NAME]) {
        const error = new Error(`Sheet "${SHEET_NAME}" not found in workbook. Available sheets: ${workbook.SheetNames.join(', ')}`);
        error.code = 'SHEET_NOT_FOUND';
        throw error;
      }

      const worksheet = workbook.Sheets[SHEET_NAME];
      const rawRows = xlsx.utils.sheet_to_json(worksheet, { defval: '' });

      // Clean & normalize rows (filter out empty rows)
      const products = rawRows
        .filter(row => {
          // Row must have at least an ID or a Name to be valid
          const id = row.ID !== undefined ? String(row.ID).trim() : '';
          const name = row.Name !== undefined ? String(row.Name).trim() : '';
          return id !== '' || name !== '';
        })
        .map(row => {
          // Normalize Price (INR)
          let price = row['Price (INR)'];
          if (typeof price === 'string') {
            price = parseFloat(price.replace(/[^0-9.-]+/g, '')) || 0;
          }

          // Normalize Stock
          let stock = row.Stock;
          if (typeof stock === 'string') {
            stock = parseInt(stock.replace(/[^0-9.-]+/g, ''), 10) || 0;
          }

          // Normalize Rating
          let rating = row.Rating;
          if (typeof rating === 'string') {
            rating = parseFloat(rating) || 0;
          }

          return {
            ID: row.ID !== undefined ? row.ID : '',
            Name: row.Name !== undefined ? String(row.Name).trim() : '',
            Category: row.Category !== undefined ? String(row.Category).trim() : '',
            'Price (INR)': typeof price === 'number' && !isNaN(price) ? price : 0,
            Stock: typeof stock === 'number' && !isNaN(stock) ? stock : 0,
            Rating: typeof rating === 'number' && !isNaN(rating) ? rating : 0,
            Status: row.Status !== undefined ? String(row.Status).trim() : 'In Stock'
          };
        });

      // Update cache
      cachedData = products;
      lastModifiedTime = stat.mtime;

      return {
        products,
        lastModified: stat.mtime,
        isCached: false
      };
    } catch (err) {
      // If file is locked (common on Windows when saving in Excel)
      const isLockError = err.code === 'EBUSY' || err.code === 'EPERM' || err.code === 'EACCES';
      attempts++;

      if (isLockError && attempts <= maxRetries) {
        // Synchronous sleep before retry
        const start = Date.now();
        while (Date.now() - start < delayMs) {
          // busy wait short interval
        }
        continue;
      }

      // If still failing but we have cached data, return cached snapshot with warning
      if (isLockError && cachedData) {
        console.warn(`[WARN] Excel file temporarily locked. Serving latest cached snapshot (${cachedData.length} items).`);
        return {
          products: cachedData,
          lastModified: lastModifiedTime,
          isCached: true,
          warning: 'File locked by Excel. Displaying latest cached snapshot.'
        };
      }

      throw err;
    }
  }
}

/**
 * GET /api/products
 * Reads products_database.xlsx from disk on every API call and returns all rows as JSON
 */
app.get('/api/products', (req, res) => {
  // Prevent any HTTP caching so every client fetch gets fresh data
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');

  try {
    const result = readExcelFile();

    if (result.lastModified) {
      res.setHeader('X-Last-Modified', result.lastModified.toISOString());
    }
    if (result.isCached) {
      res.setHeader('X-Data-Source', 'cache-locked-fallback');
    } else {
      res.setHeader('X-Data-Source', 'disk-live');
    }

    // Return all rows as JSON
    return res.status(200).json(result.products);
  } catch (error) {
    console.error('[API ERROR] Failed to read Excel database:', error.message);

    if (error.code === 'ENOENT') {
      return res.status(404).json({
        error: `Database file "products_database.xlsx" not found on disk.`,
        code: 'FILE_NOT_FOUND',
        solution: 'Make sure products_database.xlsx is placed in the project root directory or run "npm run reset-db".'
      });
    }

    if (error.code === 'SHEET_NOT_FOUND') {
      return res.status(400).json({
        error: error.message,
        code: 'SHEET_NOT_FOUND',
        solution: 'Make sure your Excel file contains a worksheet named "Products".'
      });
    }

    if (error.code === 'EBUSY' || error.code === 'EPERM' || error.code === 'EACCES') {
      return res.status(503).json({
        error: 'Excel file is currently locked by another application (e.g. Microsoft Excel).',
        code: 'FILE_LOCKED',
        solution: 'Please save changes in Excel or close the file and retry.'
      });
    }

    return res.status(500).json({
      error: 'Internal server error reading Excel file: ' + error.message,
      code: 'INTERNAL_ERROR'
    });
  }
});

/**
 * GET /api/products/:id
 * Returns a specific product by ID directly from Excel
 */
app.get('/api/products/:id', (req, res) => {
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
  try {
    const { products } = readExcelFile();
    const product = products.find(p => String(p.ID) === String(req.params.id));
    if (!product) {
      return res.status(404).json({ error: `Product with ID #${req.params.id} not found`, code: 'NOT_FOUND' });
    }
    return res.json(product);
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
});

// Store placed orders in memory
const ordersStore = [];

/**
 * POST /api/orders
 * Fully working checkout endpoint that processes and stores user orders
 */
app.post('/api/orders', (req, res) => {
  try {
    const { items, customer, paymentMethod, totals } = req.body;
    if (!items || !Array.isArray(items) || items.length === 0) {
      return res.status(400).json({ error: 'Order must contain at least one item.' });
    }
    if (!customer || !customer.fullName || !customer.address) {
      return res.status(400).json({ error: 'Please provide full name and delivery address.' });
    }

    const orderId = 'ORD-' + Math.floor(100000 + Math.random() * 900000);
    const razorpayPaymentId = req.body.razorpayPaymentId || null;
    const paymentMethodName = paymentMethod || (razorpayPaymentId ? 'Razorpay' : 'Cash on Delivery');

    const order = {
      orderId,
      createdAt: new Date().toISOString(),
      status: 'Confirmed',
      items,
      customer,
      paymentMethod: paymentMethodName,
      razorpayPaymentId: razorpayPaymentId,
      paymentStatus: razorpayPaymentId ? 'Paid & Captured' : (paymentMethodName.includes('COD') ? 'Pending (COD)' : 'Completed'),
      totals: totals || {
        subtotal: items.reduce((sum, it) => sum + (it.price * it.quantity), 0),
        shipping: 0,
        tax: 0,
        grandTotal: items.reduce((sum, it) => sum + (it.price * it.quantity), 0)
      },
      estimatedDelivery: new Date(Date.now() + 3 * 24 * 60 * 60 * 1000).toLocaleDateString('en-IN', {
        weekday: 'short',
        month: 'short',
        day: 'numeric'
      })
    };

    ordersStore.unshift(order);
    console.log(`[ORDER PLACED] ${orderId} by ${customer.fullName} for ₹${order.totals.grandTotal} via ${paymentMethodName}${razorpayPaymentId ? ' (' + razorpayPaymentId + ')' : ''}`);

    return res.status(201).json({ success: true, order });
  } catch (err) {
    return res.status(500).json({ error: 'Failed to process order: ' + err.message });
  }
});

/**
 * POST /api/reset-database
 * Allows resetting products_database.xlsx back to original seed data
 */
app.post('/api/reset-database', (req, res) => {
  try {
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

    const wb = xlsx.utils.book_new();
    const wsProducts = xlsx.utils.json_to_sheet(sampleProducts);
    const wsReadme = xlsx.utils.json_to_sheet(readmeData);

    xlsx.utils.book_append_sheet(wb, wsProducts, 'Products');
    xlsx.utils.book_append_sheet(wb, wsReadme, 'README');

    xlsx.writeFile(wb, EXCEL_FILE_PATH);

    // Invalidate local cache
    cachedData = null;

    res.json({ success: true, message: 'Database reset to original 12 products successfully.' });
  } catch (err) {
    res.status(500).json({ error: 'Failed to reset database: ' + err.message });
  }
});

// Fallback to index.html for root or client routes
app.use((req, res, next) => {
  if (req.method === 'GET' && !req.path.startsWith('/api')) {
    return res.sendFile(path.join(__dirname, 'public', 'index.html'));
  }
  next();
});

// Start Server
app.listen(PORT, () => {
  console.log('====================================================');
  console.log(`🚀 Excel Live Database Server running at:`);
  console.log(`👉 http://localhost:${PORT}`);
  console.log(`📁 Watching Excel file: ${EXCEL_FILE_PATH}`);
  console.log(`📊 Target Sheet: "${SHEET_NAME}"`);
  console.log('====================================================');
});
