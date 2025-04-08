using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
class CellData
{
    public string Cell { get; set; }
    public string Value { get; set; }
    public string Format { get; set; } = "String"; // Default
}


class Program
{
    static void Main(string[] args)
    {
        if (args.Length != 3)
        {
            Console.WriteLine("Usage: app <jsonFilePath> <excelFilePath> <sheetName>");
            return;
        }

        string jsonPath = args[0];
        string excelPath = args[1];
        string sheetName = args[2];

        if (!File.Exists(jsonPath))
        {
            Console.WriteLine("JSON file does not exist.");
            return;
        }

        var jsonText = File.ReadAllText(jsonPath);
        var cellDataList = JsonSerializer.Deserialize<List<CellData>>(jsonText);

        bool fileExists = File.Exists(excelPath);
        using var spreadsheet = fileExists
            ? SpreadsheetDocument.Open(excelPath, true)
            : SpreadsheetDocument.Create(excelPath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);

        if (!fileExists)
        {
            var workbookPart = spreadsheet.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            var sheets = spreadsheet.WorkbookPart.Workbook.AppendChild(new Sheets());
            sheets.Append(new Sheet() { Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = sheetName });
            spreadsheet.WorkbookPart.Workbook.Save();
        }

      // 🔁 Force recalculation on load (modern workaround)
        var existingCalc = spreadsheet.WorkbookPart.Workbook
        .Descendants<OpenXmlElement>()
        .FirstOrDefault(e => e.LocalName == "workbookCalcPr");

            if (existingCalc != null)
            {
                existingCalc.Remove();
            }

            var fullCalcElement = new OpenXmlUnknownElement("workbookCalcPr");
            fullCalcElement.SetAttribute(new OpenXmlAttribute("fullCalcOnLoad", null, "1"));
            spreadsheet.WorkbookPart.Workbook.Append(fullCalcElement);
            spreadsheet.WorkbookPart.Workbook.Save();


            WriteToSheet(spreadsheet, sheetName, cellDataList);
            ClearCachedFormulaValues(spreadsheet, sheetName);
            spreadsheet.WorkbookPart.Workbook.Save();
            Console.WriteLine("Done writing to Excel.");
        }

    static void ClearCachedFormulaValues(SpreadsheetDocument doc, string sheetName)
    {
        var sheet = GetSheetByName(doc, sheetName);
        if (sheet == null)
        {
            Console.WriteLine($"Sheet {sheetName} not found.");
            return;
        }

        var worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id);
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellFormula != null)
                {
                    cell.CellValue = null;
                }
            }
        }

        worksheetPart.Worksheet.Save();
    }


    static void WriteToSheet(SpreadsheetDocument doc, string sheetName, List<CellData> cellDataList)
    {
        var sheet = GetSheetByName(doc, sheetName);
        if (sheet == null)
        {
            Console.WriteLine($"Sheet {sheetName} not found.");
            return;
        }

        var worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id);
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

        foreach (var data in cellDataList)
        {
            var (column, row) = SplitCellReference(data.Cell);
            var rowElement = GetOrCreateRow(sheetData, row);
            var cell = GetOrCreateCell(rowElement, column);

            // Write value and handle format
            switch (data.Format?.ToLowerInvariant())
            {
                case "number":
                    cell.CellValue = new CellValue(data.Value);
                    cell.DataType = null; // Keep numeric format
                    break;

                case "date":
                    if (DateTime.TryParse(data.Value, out var date))
                    {
                        // Convert to Excel date serial number
                        double oaDate = date.ToOADate();
                        cell.CellValue = new CellValue(oaDate.ToString(System.Globalization.CultureInfo.InvariantCulture));
                        cell.DataType = null;
                        // Optional: Set a style index with date formatting if desired
                    }
                    else
                    {
                        Console.WriteLine($"Invalid date format: {data.Value}");
                    }
                    break;

                case "string":
                default:
                    cell.CellValue = new CellValue(data.Value);
                    cell.DataType = CellValues.String;
                    break;
            }
        }


        

        worksheetPart.Worksheet.Save();
    }

    static Sheet GetSheetByName(SpreadsheetDocument doc, string name)
    {
        foreach (Sheet sheet in doc.WorkbookPart.Workbook.Sheets)
        {
            if (sheet.Name == name) return sheet;
        }
        return null;
    }

    static Row GetOrCreateRow(SheetData sheetData, uint rowIndex)
    {
        var row = sheetData.Elements<Row>()?.FirstOrDefault(r => r.RowIndex == rowIndex);
        if (row == null)
        {
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
        }
        return row;
    }

   static Cell GetOrCreateCell(Row row, string columnName)
{
    var refId = columnName + row.RowIndex;
    var existingCell = row.Elements<Cell>()?.FirstOrDefault(c => c.CellReference == refId);
    if (existingCell != null)
    {
        return existingCell;
    }

    // New cell — but let's find the style index from a nearby existing cell if possible
    var cell = new Cell() { CellReference = refId };

    // Optional: Try to find a "template" cell in the same column with a StyleIndex
    var templateCell = row.Elements<Cell>()
        .FirstOrDefault(c => GetColumnName(c.CellReference?.Value) == columnName && c.StyleIndex != null);

    if (templateCell != null)
    {
        cell.StyleIndex = templateCell.StyleIndex;
    }

    row.Append(cell);
    return cell;
}

static string GetColumnName(string cellReference)
{
    return new string(cellReference?.TakeWhile(char.IsLetter).ToArray());
}


    static (string column, uint row) SplitCellReference(string reference)
    {
        string column = new string(reference.TakeWhile(char.IsLetter).ToArray());
        string rowStr = new string(reference.SkipWhile(char.IsLetter).ToArray());
        return (column, uint.Parse(rowStr));
    }
}
