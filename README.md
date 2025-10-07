# 🧾 Quattro Revisions VBA Macro

[![Excel](https://img.shields.io/badge/Excel-VBA-217346?logo=microsoft-excel&logoColor=white)](https://learn.microsoft.com/en-us/office/vba/api/overview/)
[![Platform](https://img.shields.io/badge/Platform-Windows-blue)]()
[![License](https://img.shields.io/badge/License-MIT-lightgrey.svg)]()

---

## 📋 Overview
**Quattro Revisions** is a VBA automation tool that extracts **revision identifiers** from web pages linked within an Excel worksheet.  
The macro opens each hyperlink, reads the browser window title, searches for `"Rev"`, and records the detected revision (e.g., `RevA2`) next to the corresponding link.

It then performs cleanup, logs the date/time, and saves a **macro-free copy** of the workbook for reporting or archival.

---

## ✨ Key Features

- 🔗 Opens hyperlinks in specific cell ranges:
  - `A3:A32`
  - `G7:G8`
  - `G12:G22`
- ⏱ Waits for browser load before reading the window title.
- 🧩 Detects “Rev” and captures up to **4 characters** following it.
- 🗒 Writes extracted revisions beside each hyperlink.
- 🧼 Cleans up automatically:
  - Removes hyperlinks (keeping visible text & cell formatting).
  - Reapplies **All Borders** to affected ranges.
  - Unmerges and clears **cell G30**.
  - Deletes any **Form Control buttons**.
- 🕒 Records completion timestamp:
  - **Date → `K34`**
  - **Time → `K35`**
- 💾 Saves a clean, macro-free workbook:

🧠 How It Works

The macro loops through all defined hyperlink ranges.

Each link opens in your default browser (Chrome, Edge, or Firefox).

The macro searches active browser windows for "Rev".

The detected revision is extracted and recorded beside the link.

When all links are processed:

Hyperlinks are removed (values retained).

Formatting is restored.

Cell G30 is unmerged and cleared.

Any Form Control buttons are deleted.

The workbook is saved as a .xlsx file with no macros.

🖥️ Requirements
Component	Requirement
Excel Version	Microsoft Excel 2016 or later
OS	Windows (API-dependent)
Libraries Used	user32.dll, kernel32.dll (for window management and Sleep)
Browsers Supported	Chrome, Edge, Firefox, Internet Explorer (legacy)
🧩 API Calls Used
API Function	Purpose
FindWindow	Finds the top-level browser window
FindWindowEx	Iterates through multiple browser windows
GetWindowText	Reads the browser’s title bar text
Sleep	Pauses execution while pages load

All declarations are PtrSafe for 64-bit compatibility.

🚀 Usage

Open the workbook containing the macro.

Ensure all hyperlinks are valid and point to the desired documents.

Run the macro:

In Excel:
Developer → Macros → Select ExtractChromeTabRev → Run

Allow browser tabs to open — the macro will handle each automatically.

When complete:

Revisions appear beside each hyperlink.

A summary message confirms completion.

The file Quattro Revisions.xlsx is saved in the same directory.

📁 Output Example
Link (A)	Extracted Revision (B)
Drawing123.pdf
	RevA
Layout456.pdf
	RevB1
Assembly789.pdf
	RevNotFound
🧰 Troubleshooting
Issue	Possible Cause	Solution
NoWindow in result	Browser window title not detected	Ensure Chrome/Edge is active and not minimized
RevNotFound	Page doesn’t include “Rev” in title	Verify the linked file uses “Rev” convention
Borders disappear	Formatting removed by hyperlink deletion	Handled automatically in current version
Macro doesn’t run	Macros disabled	Enable macros via File → Options → Trust Center
🧑‍💻 Developer Notes

The macro uses Sleep to wait for page loads — this can be adjusted for slower network connections.

Compatible with both 32-bit and 64-bit Excel installations.

For repeatability, it’s recommended to keep the hyperlink layout and range definitions consistent.

📜 License

This project is licensed under the MIT License
.
You are free to modify, distribute, and use it for personal or commercial purposes.

👤 Author

Quattro Revisions Macro
Developed by J. S.
© 2025 — Internal Automation Utility

⭐ If you find this useful, please star the repository!
