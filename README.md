# DocSmart Office Tools - Excel to PDF Converter

A modern, fast, and 100% private in-browser **Excel to PDF Converter** web application. Built with clean HTML5, CSS3, and modern client-side JavaScript.

---

## 🚀 GitHub Pages Readiness Status: READY ✅

This project has been thoroughly inspected and prepared for free **GitHub Pages** hosting:
- ✅ **Fully Static**: 100% client-side code (HTML5, CSS3, ES6 JavaScript).
- ✅ **Zero Backend Dependencies**: No Node.js, Python, PHP, or database required.
- ✅ **Works from Root**: `index.html` resides at the repository root.
- ✅ **Relative Paths**: All CSS, JavaScript, and library paths use relative URLs, ensuring compatibility with GitHub Pages repository subpaths (e.g. `https://<username>.github.io/<repo-name>/`).
- ✅ **100% In-Browser Conversion**: Excel parsing and PDF generation run entirely in the browser using standalone libraries (`SheetJS` and `jsPDF + AutoTable`).
- ✅ **Fully Offline-Capable**: All vendor libraries are stored locally in `libs/` with automatic CDN fallbacks.

---

## 🌟 Highlights & Features

- **Drag & Drop Upload**: Seamlessly drop `.xlsx`, `.xls`, or `.csv` files or browse from your computer.
- **100% Client-Side & Private**: Your spreadsheet is parsed and converted directly inside your browser. No files are uploaded to any server, guaranteeing total data security and privacy.
- **Interactive Multi-Sheet Support**: Detects multiple sheets in your workbook. You can convert all sheets into one multi-page PDF or choose an individual sheet.
- **Spreadsheet Live Preview**: Inspect table rows and columns directly on the page before converting.
- **Customization Options**:
  - **Orientation**: Auto-Detect (smart column count heuristic), Portrait, or Landscape.
  - **Page Sizes**: A4, US Letter, and US Legal.
  - **Table Themes**: Excel Emerald Green, Corporate Navy, Clean Minimalist, or Monochrome (B&W).
  - **Font Scaling**: Compact, Normal, or Large.
  - **Headers & Footers**: Customizable document title, timestamp, and automatic "Page X of Y" pagination.
- **In-Browser PDF Preview**: View the generated PDF right in the app before downloading.
- **1-Click Download & Print**: Download the converted PDF or send it directly to your printer.
- **Zero Installation & Offline Ready**: All open-source libraries (`SheetJS`, `jsPDF`, and `jspdf-autotable`) are saved locally in the `libs/` folder. It works completely offline without an internet connection.
- **Built-in Sample Generator**: Includes a "Try with Sample Data" button so you can test the app instantly without even needing an Excel file on hand.

---

## 🌐 How to Publish to GitHub Pages (Free Hosting)

Follow these simple steps to host your application for free on GitHub Pages:

### Step 1: Create a New GitHub Repository
1. Log in to your account at [GitHub.com](https://github.com).
2. Click the **+** (plus icon) in the top right corner and choose **New repository**.
3. Name your repository (e.g., `docsmart-office-tools` or `excel-to-pdf`).
4. Set visibility to **Public** (required for free GitHub Pages accounts).
5. Leave "Add a README file" unchecked (you already have this one).
6. Click **Create repository**.

---

### Step 2: Upload Project Files to GitHub

Choose whichever method you prefer:

#### Option A: Web Browser Upload (Easiest - No Software to Install)
1. On your newly created repository page, click the link **uploading an existing file**.
2. Select or drag-and-drop the files and folders from this project folder:
   - `index.html` *(must be in the root directory)*
   - `README.md`
   - `css/` (containing `styles.css`)
   - `js/` (containing `app.js` and `converter.js`)
   - `libs/` (containing `xlsx.full.min.js`, `jspdf.umd.min.js`, `jspdf.plugin.autotable.min.js`)
   - `sample/` (containing `sample_company_sales.csv`)
   - `start.bat` *(optional local launcher)*
   - `serve.ps1` *(optional local server)*
3. In the "Commit changes" box at the bottom, enter a summary like `Initial release for GitHub Pages`.
4. Click **Commit changes**.

#### Option B: Using Git Command Line
If you have Git installed, open PowerShell or Terminal in this folder and run:
```bash
git init
git add .
git commit -m "Initial release for GitHub Pages"
git branch -M main
git remote add origin https://github.com/<YOUR_GITHUB_USERNAME>/<YOUR_REPOSITORY_NAME>.git
git push -u origin main
```

#### Option C: Using GitHub Desktop
1. Open GitHub Desktop.
2. Go to **File** -> **Add Local Repository...** and select this project folder.
3. Publish repository to GitHub.

---

### Step 3: Enable GitHub Pages in Repository Settings
1. On your GitHub repository page, click the **Settings** tab (the gear icon near top right).
2. In the left navigation menu, click **Pages** (under the "Code and automation" section).
3. Under **Build and deployment**:
   - **Source**: Select **Deploy from a branch**.
   - **Branch**: Select `main` (or `master`).
   - **Folder**: Select `/ (root)`.
4. Click **Save**.
5. Wait 1 to 2 minutes for GitHub Actions to build and deploy your site.
6. Refresh the Settings > Pages page. You will see a banner at the top:
   > **"Your site is live at `https://<YOUR_GITHUB_USERNAME>.github.io/<YOUR_REPOSITORY_NAME>/`"**

---

### Step 4: Verify Your Live Site
1. Click the URL provided by GitHub.
2. Your **DocSmart Office Tools** application will load instantly with HTTPS encryption.
3. Test by clicking **Try with Sample Data**, selecting your desired options, and clicking **Convert to PDF Now**.
4. Test uploading your own `.xlsx` or `.csv` files.

---

## 💻 How to Run Locally on Your Computer (Offline)

You can run this project locally on your machine at any time without an internet connection:

### Option 1: Double-Click `index.html` (Instant)
1. Open this project folder in Windows File Explorer.
2. Double-click **`index.html`**.
3. It will open directly in Edge, Chrome, Firefox, or Safari.

### Option 2: Double-Click `start.bat` (Windows Shortcut)
1. Double-click the file named **`start.bat`**.
2. It automatically launches the app in your default web browser.

### Option 3: Local HTTP Preview (`serve.ps1`)
If you want to simulate an HTTP web server locally:
1. Right-click `serve.ps1` and choose **Run with PowerShell** (or run `powershell -ExecutionPolicy Bypass -File .\serve.ps1`).
2. Open `http://localhost:8085` in your browser.

---

## 📂 Project Structure

```
Doc Smart Office Tools/
├── index.html                        # Website root entry point (HTML5 structure)
├── css/
│   └── styles.css                    # Responsive stylesheet (100% self-contained)
├── js/
│   ├── app.js                        # UI controller (drag & drop, tabs, preview, events)
│   └── converter.js                  # Engine that parses Excel and generates PDFs
├── libs/                             # Standalone client-side libraries (vendor scripts)
│   ├── xlsx.full.min.js              # SheetJS (v0.18.5): Excel/CSV parser
│   ├── jspdf.umd.min.js              # jsPDF (v2.5.1): PDF document builder
│   └── jspdf.plugin.autotable.min.js # AutoTable (v3.8.2): Table layout & pagination
├── sample/
│   └── sample_company_sales.csv      # Sample spreadsheet for testing
├── start.bat                         # Quick launcher for local Windows use
├── serve.ps1                         # Optional local HTTP preview server
└── README.md                         # Documentation & GitHub Pages publishing guide
```

---

## 💡 Technical Architecture

- **Static Front-End**: Pure client-side web technology. Compatible with all modern browsers (Chrome, Edge, Firefox, Safari, Opera).
- **Excel Parsing**: Handled via `XLSX.read(arrayBuffer)` directly inside `converter.js`. Extracts sheet names, row cells, and column lengths in memory.
- **PDF Generation**: Handled via `new jsPDF(...)` and `.autoTable(...)`. Calculates column widths, wraps text, generates header rows, alternating zebra stripes, date timestamps, and "Page X of Y" footers.
- **Zero Remote Latency**: Because conversion occurs in the browser's JavaScript execution thread, processing is instantaneous for small and medium workbooks with zero network latency.

---

## 🔒 Privacy & Security

This application runs **100% client-side in your web browser**. 
- No files are ever uploaded or transmitted across the network.
- Zero server-side storage or tracking.
- Compliant with sensitive corporate data protection requirements.
- Works in air-gapped and completely offline environments.
