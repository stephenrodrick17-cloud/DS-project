# NexusStore — Live Excel-Driven E-Commerce with Razorpay Integration

[![Live Excel Sync](https://img.shields.io/badge/Database-Excel%20XLSX-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://github.com/stephenrodrick17-cloud/DS-project)
[![Node.js](https://img.shields.io/badge/Backend-Node.js%20%2B%20Express-339933?style=for-the-badge&logo=nodedotjs&logoColor=white)](https://github.com/stephenrodrick17-cloud/DS-project)
[![Payment Gateway](https://img.shields.io/badge/Payment-Razorpay%20Checkout-0C2340?style=for-the-badge&logo=razorpay&logoColor=white)](https://github.com/stephenrodrick17-cloud/DS-project)
[![Frontend](https://img.shields.io/badge/Frontend-Vanilla%20HTML%2FCSS%2FJS-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)](https://github.com/stephenrodrick17-cloud/DS-project)

A production-grade, multi-page e-commerce web application where a local **Microsoft Excel workbook (`products_database.xlsx`) acts as the single source of truth database**.

When an administrator edits a cell in the Excel spreadsheet and saves it (**Ctrl + S**), the website automatically detects the changes within **5 seconds** without needing a server restart or manual page reload. Any modified value lights up with a glowing **golden pulse animation**, and customer orders are processed through an interactive **Razorpay Payment Gateway** simulation.

---

## 📑 Table of Contents
- [1. Overview & Core Philosophy](#1-overview--core-philosophy)
- [2. System Architecture](#2-system-architecture)
- [3. Complete Folder Structure](#3-complete-folder-structure)
- [4. How It Works (Detailed Component Breakdown)](#4-how-it-works-detailed-component-breakdown)
  - [A. Backend Engine (`server.js`)](#a-backend-engine-serverjs)
  - [B. Excel Database (`products_database.xlsx`)](#b-excel-database-products_databasexlsx)
  - [C. Frontend Architecture (`app.js`, `index.html`, `style.css`)](#c-frontend-architecture-appjs-indexhtml-stylecss)
  - [D. Razorpay Payment Gateway Integration](#d-razorpay-payment-gateway-integration)
- [5. Step-by-Step Setup & Running Guide](#5-step-by-step-setup--running-guide)
- [6. Testing Guide (Live Excel & Payment Walkthrough)](#6-testing-guide-live-excel--payment-walkthrough)
- [7. API Reference](#7-api-reference)
- [8. Troubleshooting & FAQ](#8-troubleshooting--faq)

---

## 1. Overview & Core Philosophy

In traditional web development, databases require complex database engines (PostgreSQL, MongoDB, MySQL). However, in retail, inventory management, and small businesses, non-technical managers frequently maintain inventory, prices, and catalogs directly in spreadsheets.

**NexusStore bridges this gap**:
1. **Zero Database Configuration**: No SQL server or daemon required. The file `products_database.xlsx` is the database.
2. **True Live Synchronization**: Every 5 seconds, the website re-queries the file from disk via Express and SheetJS.
3. **Visual Change Attribution**: When a price drops or stock depletes in Excel, the frontend visually highlights the changed value with a smooth keyframe flash animation.
4. **End-to-End E-Commerce Workflow**: Complete with Product Detail Pages (PDP), Cart, Promo Codes, multi-option Checkout, an interactive **Razorpay Gateway demo**, and verified Order Receipts.

---

## 2. System Architecture

```text
 ┌────────────────────────────────────────────────────────┐
 │            Microsoft Excel / LibreOffice               │
 │            (products_database.xlsx)                   │
 └──────────────────────────┬─────────────────────────────┘
                            │ (User edits & saves with Ctrl + S)
                            ▼
 ┌────────────────────────────────────────────────────────┐
 │           Node.js + Express Server (server.js)         │
 │  - Reads Excel buffer on disk on every request        │
 │  - Graceful Windows file-lock retry (EBUSY / EPERM)    │
 │  - Serves static assets & REST endpoints:              │
 │    * GET /api/products                                 │
 │    * GET /api/products/:id                             │
 │    * POST /api/orders                                  │
 │    * POST /api/reset-database                          │
 │    * GET /api/meta                                     │
 └──────────────────────────┬─────────────────────────────┘
                            │ (JSON API over HTTP)
                            ▼
 ┌────────────────────────────────────────────────────────┐
 │       Vanilla JS Client Application (public/app.js)    │
 │  - Client-side Hash Router:                            │
 │    #catalog | #product/:id | #cart | #checkout |       │
 │    #order-success | #sync-status                       │
 │  - 5-second interval heartbeat & diff detection engine │
 │  - Persistent localStorage Shopping Cart               │
 │  - Interactive Razorpay Checkout Modal (UPI/Cards/etc) │
 └────────────────────────────────────────────────────────┘
```

---

## 3. Complete Folder Structure

```text
DS-project/
├── .gitignore               # Ignores node_modules, logs, and OS caches
├── package.json             # Project metadata, dependencies, and npm scripts
├── README.md                # Comprehensive system documentation (this file)
├── reset_database.js        # Standalone utility to regenerate/seed sample Excel data
├── server.js                # Express backend reading Excel database on every request
├── products_database.xlsx   # The live single source of truth database file
└── public/                  # Frontend single-page / multi-view static application
    ├── index.html           # Semantic HTML5 layout, views, drawers, and modal dialogs
    ├── style.css            # Custom responsive stylesheet with glassmorphism & dark aesthetics
    └── app.js               # Router, auto-polling engine, cart, Razorpay flow, and DOM logic
```

---

## 4. How It Works (Detailed Component Breakdown)

### A. Backend Engine (`server.js`)
- **Reading from Disk**: Rather than keeping data in memory or caching the workbook, the server uses `fs.readFileSync(EXCEL_FILE_PATH)` and `xlsx.read(buffer, { type: 'buffer' })` on every call to `/api/products`.
- **Handling Windows File Locks**: On Windows, when Microsoft Excel saves a file, it momentarily places an exclusive write lock (`EBUSY` / `EPERM`). The server implements an automated retry mechanism:
  - If locked, it pauses briefly (120ms) and retries up to 2 times.
  - If still locked, it serves the last successful snapshot with an `X-Data-Source: cache-locked-fallback` header, preventing browser error crashes while Excel finishes writing.
- **Data Normalization**: Cleans cell values, strips unwanted characters, parses numbers, filters blank rows, and structures JSON output cleanly:
  ```json
  {
    "ID": 2,
    "Name": "Mechanical Keyboard",
    "Category": "Electronics",
    "Price (INR)": 2499,
    "Stock": 5,
    "Rating": 4.6,
    "Status": "Low Stock"
  }
  ```

### B. Excel Database (`products_database.xlsx`)
- **Target Worksheet**: `"Products"`
- **Required Columns**:
  1. `ID` (Unique identifier for each item)
  2. `Name` (Product name)
  3. `Category` (Electronics, Clothing, Footwear, Accessories, Home, Stationery)
  4. `Price (INR)` (Numerical retail price in Indian Rupees)
  5. `Stock` (Inventory quantity available in warehouse)
  6. `Rating` (Customer rating score out of 5)
  7. `Status` (`In Stock`, `Low Stock`, or `Out of Stock`)

### C. Frontend Architecture (`app.js`, `index.html`, `style.css`)
- **Client-Side Hash Router**: Provides a snappy multi-page user experience without full page reloads:
  - `#catalog`: Main shop catalog with search, category filter chips, sorting, and Grid/Table view.
  - `#product/:id`: Dedicated Product Detail Page (PDP) with high-res photography, warranty perks, live stock bar, and quantity selector.
  - `#cart`: Shopping cart overview with live totals, tax calculation, and coupon input.
  - `#checkout`: Shipping details form and payment method selection.
  - `#order-success`: Purchase receipt, Order ID, and Razorpay transaction summary.
  - `#sync-status`: Live diagnostic dashboard showing Excel file stats and testing guide.
- **5-Second Heartbeat Engine**:
  - Automatically fetches `/api/products` every 5 seconds.
  - Keeps a visual countdown badge (`5s` ... `1s`) in the top navigation bar.
- **Field-by-Field Diff Detector**:
  - Compares the incoming JSON array against `previousProductsMap`.
  - Pinpoints exactly which field changed (`Price (INR)`, `Stock`, `Status`, `Rating`, or `Name`).
  - Injects CSS class `.value-changed` to trigger a golden glow pulse animation (`highlightFlash`) that lasts 2 seconds.
  - Displays a toast alert: `⚡ Updated in Excel: Mechanical Keyboard (Price: ₹2,999 ➔ ₹2,499)`.
- **Dynamic Product Photography**:
  - Features curated, high-resolution photography for all catalog items.
  - Includes smart category-based fallbacks for any new product added directly to Excel.

### D. Razorpay Payment Gateway Integration
When customers proceed to checkout with the recommended **Razorpay Secure** payment option, the application launches an authentic **Razorpay Checkout Modal**:
1. **Razorpay Header**: Displays official Razorpay blue branding, NexusStore merchant title, order preview ID, and amount payable.
2. **Sandbox Notice**: Prominently highlights **`[TEST MODE]`** so users know no real money is deducted.
3. **Four Native Payment Tabs**:
   - ⚡ **UPI & QR Code**: Live SVG QR Code with glowing scanning radar, validity countdown timer (`04:59`), and UPI ID verification.
   - 💳 **Cards**: Card number field with Visa/Mastercard detection, expiration date, and CVV inputs.
   - 🏦 **Netbanking**: Direct selector for leading banks (HDFC Bank, SBI, ICICI Bank, Axis Bank, Kotak, PNB).
   - 👛 **Wallets**: PhonePe, Paytm, and Mobikwik.
4. **Interactive Sandbox Controls**:
   - `✅ Simulate Success (Instant Pay)`: Emulates successful 256-bit bank authorization, generates a real-format Razorpay Payment ID (e.g. `pay_B8KM2LQ9XP`), marks order as **`Paid & Captured`**, and routes to the receipt.
   - `❌ Simulate Bank Decline`: Tests error handling by displaying a simulated bank decline error code (`1002 - Bank authorization declined`).

---

## 5. Step-by-Step Setup & Running Guide

### Prerequisites
- **Node.js** (v18.x or higher)
- **npm** (v9.x or higher)
- **Git**

### Installation

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/stephenrodrick17-cloud/DS-project.git
   cd DS-project
   ```

2. **Install Node Dependencies**:
   ```bash
   npm install
   ```

3. **Start the Express Application**:
   ```bash
   node server.js
   ```
   *(Or run in watch mode for development: `npm run dev`)*

4. **Open in Browser**:
   Open your browser and navigate to:
   👉 **`http://localhost:3000`**

---

## 6. Testing Guide (Live Excel & Payment Walkthrough)

### Test A: Live Excel Price Edit
1. Open `http://localhost:3000` in your web browser.
2. Open `products_database.xlsx` in **Microsoft Excel**, **LibreOffice**, or **VS Code**.
3. In the **"Products"** sheet, change row 3 (`Mechanical Keyboard`):
   - Set `Price (INR)` from `2499` to `1799`.
4. Press **Ctrl + S** to save the Excel file.
5. Watch the website within **5 seconds**: the price automatically animates with a glowing golden pulse, updates to **₹1,799**, and pops a live toast notification!

### Test B: Live Out of Stock Behavior
1. In `products_database.xlsx`, find row 2 (`Wireless Mouse`).
2. Change `Stock` to `0` and `Status` to `Out of Stock`. Save (**Ctrl + S**).
3. The badge turns **Red**, stock becomes `0`, and on the product detail page the **Add to Cart** and **Buy Now** buttons are automatically disabled with an "Out of Stock" state.

### Test C: Adding a New Product in Excel
1. Add a new row at the bottom of the "Products" sheet:
   - `ID`: `14`
   - `Name`: `Active Noise-Cancelling Headphones`
   - `Category`: `Electronics`
   - `Price (INR)`: `6499`
   - `Stock`: `15`
   - `Rating`: `4.9`
   - `Status`: `In Stock`
2. Save (**Ctrl + S**).
3. Within 5 seconds, a new product card appears on the site with photography, category badge, and active purchasing options!

### Test D: Razorpay Checkout Demo
1. Browse products and click **"Add to Cart"**.
2. Open the cart drawer and click **"Checkout Now ->"**.
3. On the Checkout page, ensure **"Razorpay Secure"** is selected.
4. Click **"Confirm & Place Order"**.
5. The **Razorpay Checkout Dialog** will open. Explore the tabs (UPI QR, Cards, Netbanking, Wallets).
6. Click **"✅ Simulate Success (Instant Pay)"**.
7. The bank authorization spinner will run for ~1.2s and redirect you to the **Order Success Page** showing your **Order ID**, **Razorpay Payment ID (`pay_...`)**, and verified purchase receipt!

---

## 7. API Reference

| Method | Endpoint | Description |
| :--- | :--- | :--- |
| `GET` | `/api/products` | Re-reads `products_database.xlsx` from disk and returns all active products as JSON. |
| `GET` | `/api/products/:id` | Returns details for a single product by ID directly from the Excel sheet. |
| `POST` | `/api/orders` | Accepts cart items, customer details, and Razorpay payment ID, and generates a confirmed order receipt. |
| `POST` | `/api/reset-database` | Resets `products_database.xlsx` back to its original default 12 products. |
| `GET` | `/api/meta` | Diagnostic endpoint returning database file size, row count, and OS last modified timestamp. |

---

## 8. Troubleshooting & FAQ

- **Q: Excel displays "File in Use" or website returns error when saving?**  
  *A: The server includes automated retries for transient Windows locks (`EBUSY`). If Excel holds an exclusive lock for an extended period, the server seamlessly serves the latest cached snapshot.*

- **Q: How do I restore the original 12 products if I corrupt my Excel file?**  
  *A: Run `npm run reset-db` from the terminal or click the "Reset Database" button on the `#sync-status` page.*

- **Q: Can I use promo codes?**  
  *A: Yes! Enter promo code **`EXCEL10`** in the shopping cart to receive an instant 10% discount.*

---

## 👨‍💻 Author & License
- **Author**: Stephen Rodrick ([@stephenrodrick17-cloud](https://github.com/stephenrodrick17-cloud))
- **Repository**: [https://github.com/stephenrodrick17-cloud/DS-project](https://github.com/stephenrodrick17-cloud/DS-project)
- **License**: MIT
