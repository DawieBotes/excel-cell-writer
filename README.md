# Excel Cell Writer CLI

A lightweight C# .NET command-line application to update specific Excel cells based on a JSON input file, using the **OpenXML SDK**.

Supports automatic recalculation of formulas on file open, and respects existing cell formatting (e.g., currency, date, etc.).

---

## 🔧 Features

- Writes cell values to a given Excel sheet from JSON  
- Preserves existing formatting (e.g., currency styles)  
- Supports different data formats: String, Number, Date  
- Triggers full Excel recalculation on next open  
- Clears cached formula results to ensure updated output  
- Automatically creates a new Excel file or sheet if missing  

---

## 📦 Requirements

- .NET 6 or newer  
- OpenXML SDK  

## To use:

Download the release folder.

Run the EXE:

```
excel-cell-writer.exe -- <jsonFilePath> <excelFilePath> <sheetName>
```

Or Install with:

```
dotnet add package DocumentFormat.OpenXml
```

---

## 🚀 Usage

```
dotnet run -- <jsonFilePath> <excelFilePath> <sheetName>
```

### Arguments

| Argument        | Description                          |
|-----------------|--------------------------------------|
| `jsonFilePath`  | Path to JSON file with cell data     |
| `excelFilePath` | Path to Excel file (existing or new) |
| `sheetName`     | Sheet to write data into             |

---

## 📄 JSON Format

```json
[
  { "Cell": "A1", "Value": "Hello", "Format": "String" },
  { "Cell": "B2", "Value": "123.45", "Format": "Number" },
  { "Cell": "C3", "Value": "2025-04-08", "Format": "Date" }
]
```

- `Cell`: Target cell reference (e.g., "A1", "B5")  
- `Value`: The value to write  
- `Format`: Optional format type — `"String"` (default), `"Number"`, or `"Date"`  

---

## 📁 Example

**Sample JSON:**

```json
[
  { "Cell": "D4", "Value": "500", "Format": "Number" },
  { "Cell": "D5", "Value": "1000", "Format": "Number" },
  { "Cell": "D6", "Value": "=SUM(D4:D5)", "Format": "String" }
]
```

**Command:**

```
dotnet run -- ./data/input.json ./output/report.xlsx "Sheet1"
```

---

## 📓 Notes

- If the Excel file doesn't exist, it will be created with the given sheet.  
- Recalculation is forced by:
  - Setting `fullCalcOnLoad = true`
  - Clearing cached values of all formula cells  
- Date values are written as Excel serial numbers and will appear correctly if the cell has date formatting.

---

## 🧪 Development & Debugging

If you're working from Visual Studio Code:

1. Ensure you have the C# extension or Dev Kit installed  
2. Build and run:

```
dotnet build
dotnet run -- <args>
```

---

## 📝 License

MIT — free to use and modify.
