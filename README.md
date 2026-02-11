![GitHub last commit](https://img.shields.io/github/last-commit/webdev-jason/Quattro-Revision-Checker?style=flat-square)
![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=flat-square&logo=microsoftexcel&logoColor=white)
![VBA](https://img.shields.io/badge/VBA-007ACC?style=flat-square&logo=visualbasic&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D4?style=flat-square&logo=windows&logoColor=white)
![GitHub license](https://img.shields.io/github/license/webdev-jason/Quattro-Revision-Checker?style=flat-square)

# 📑 Quattro Revision Checker

An automation utility for Microsoft Excel that streamlines document revision tracking. This VBA macro automatically opens hyperlinked documents in a web browser, extracts the revision level from the browser tab title, and records the data directly into the active worksheet.

## ✨ Key Features
* **🌐 Browser Integration**: Automatically iterates through defined cell ranges and opens links in the default system browser (Chrome, Edge, or Firefox).
* **🔍 Intelligent Extraction**: Dynamically searches active window titles for "Rev" identifiers and captures up to 4 subsequent characters.
* **🧹 Automated Cleanup**:
    * Removes active hyperlinks while preserving text values and cell formatting.
    * Re-applies borders and gridlines to affected ranges for a clean final report.
    * Deletes temporary form controls and UI buttons used during the process.
* **📁 Macro-Free Output**: Automatically saves a copy of the processed workbook as a standard `.xlsx` file for distribution or archiving.
* **🕒 Audit Logging**: Records completion timestamps (Date and Time) directly into designated summary cells.

## 🛠️ Requirements & Technical Details
* **Host**: Microsoft Excel 2016 or later (Windows-based).
* **Dependencies**: Uses `user32.dll` and `kernel32.dll` API calls for window management and execution pauses.
* **Compatibility**: Fully compatible with both 32-bit and 64-bit Excel installations via `PtrSafe` declarations.

## 🚀 Usage

1. **Prepare**: Open the workbook and ensure your hyperlinks are located in the defined ranges (A3:A32, G7:G8, G12:G22).
2. **Execute**: Navigate to `Developer` → `Macros` → `ExtractChromeTabRev` and click **Run**.
3. **Review**: The macro will handle tab switching and extraction automatically. Once complete, a summary message will appear and a cleaned file named `Quattro Revisions.xlsx` will be generated.

## 👤 Author
**Jason Sparks** - [GitHub Profile](https://github.com/webdev-jason)

## 📄 License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
