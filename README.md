# 🔄 AutoReplace\_TextWithCode\_OnChange

Automatically replace user-entered text with standardized codes based on a lookup table.

---

## 📌 Overview

This Excel VBA script auto-replaces text values entered into a specific range (e.g., names, job titles, countries) with corresponding codes from a lookup list. It's useful for:

* ✅ Converting country names to ISO codes
* ✅ Mapping job titles to job codes
* ✅ Translating user inputs to database-ready values
* ✅ Automatically standardizing pasted data

---

## ⚙️ How It Works

When the user types or pastes something into a specified range (e.g., columns A to R):

1. The macro checks if the value exists in a lookup table.
2. If a match is found, it automatically replaces the input with the corresponding code.

The lookup table should be on a separate sheet (e.g., a sheet named `Lists`) with:

* **Column A**: The original labels (e.g., names, job titles)
* **Column B**: The corresponding codes or IDs

---

## 📂 Setup Instructions

1. **Open your Excel workbook.**
2. Press **Alt + F11** to open the **Visual Basic for Applications (VBA)** editor.
3. In the **Project Explorer** window, find the sheet where you want this automation to work.
4. **Double-click that specific worksheet name** (e.g., `Sheet1 (DataEntry)`) — this is very important!

   * 🟡 *This script must go into the **worksheet module** (not a general module)* because it's triggered by a change event (`Worksheet_Change`).
5. **Paste the VBA code** into the code window for that sheet.
6. ✅ **Customize these values in the script to fit your needs:**

   | Variable             | Purpose                                         | Default        |
   | -------------------- | ----------------------------------------------- | -------------- |
   | `Me.Range("A:R")`    | The input range to monitor for changes          | Columns A to R |
   | `lookupSheetName`    | The name of the sheet containing your lookup    | "Lists"        |
   | `lookupRangeAddress` | The two-column range for lookup (label to code) | "A\:B"         |

   🔧 Adjust these three lines in the code to match your sheet structure.

---

## 🔍 Example Use Case

If a user types `Engineer` into cell B5, and the lookup table has:

```
A          | B
-----------|----------
Engineer   | ENG001
Manager    | MGR002
```

The macro will replace `Engineer` with `ENG001` automatically.

---

## 🧠 Advanced Notes

* The script ignores cells that aren't in the specified range.
* It suppresses errors silently (e.g., if the value isn't found, it leaves it untouched).
* It disables `Application.EnableEvents` temporarily to avoid recursion.

---

## 🚀 Ideal For:

* HR systems
* Inventory or code mapping
* Data cleaning
* Prepping Excel data for import into a database or ERP system

---

## 📄 License

MIT License — use freely, contribute back if helpful 💙

---

## 👏 Author

Created by Mohamed El-ansary. This tool was built to help with structured data transformations in Excel workflows.

---
