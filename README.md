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
2. **Right-click the sheet tab** where users will input data → `View Code`.
3. **Paste the VBA code** into the worksheet module.
4. Adjust these settings in the script:

   * `Me.Range("A:R")` → the range to monitor
   * `lookupSheetName` → name of the sheet with your lookup list
   * `lookupRangeAddress` → e.g., "A\:B" for two-column lookup

No buttons or forms needed. It runs automatically when the user edits a cell in the range.

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
