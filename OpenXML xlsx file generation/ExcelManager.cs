using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXML_xlsx_file_generation
{
    internal class ExcelManager : IDisposable
    {
        private SpreadsheetDocument spreadsheetDocument;
        private WorkbookPart workbookpart;
        private WorksheetPart worksheetPart;
        private SheetData sheetData;
        private int lastLineNumber;

        public ExcelManager() { }

        /// <summary>
        /// Create a new instance of ExcelManager and create a new file.
        /// </summary>
        /// <param name="fileFullPath">full path of the file (folder + file name + extension)</param>
        /// <param name="sheetName">name of the sheet to add, null not to add any</param>
        /// <remarks>use CreateSheet to create a new sheet if you don't provide sheetName</remarks>
        public ExcelManager(string fileFullPath, string sheetName = null)
        {
            CreateFile(fileFullPath, sheetName);
        }

        /// <summary>
        /// Create a new instance of ExcelManager and create a new file.
        /// </summary>
        /// <param name="folder">folder that will receive the file</param>
        /// <param name="fileName">name of the file including its extension</param>
        /// <param name="sheetName">name of the sheet to add, null not to add any</param>
        public ExcelManager(string folder, string fileName, string sheetName = null)
        {
            CreateFile(Path.Combine(folder, fileName), sheetName);
        }

        public void Dispose()
        {
            Close();
            spreadsheetDocument.Dispose();
        }

        /// <summary>
        /// Create a new file.
        /// </summary>
        /// <param name="path">full path to the file to create</param>
        /// <param name="sheetName">name of the sheet to add, null not to add any</param>
        /// <remarks>use CreateSheet to create a new sheet if you don't provide sheetName</remarks>
        public void CreateFile(string path, string sheetName = null)
        {
            spreadsheetDocument = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);

            workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            if (!string.IsNullOrEmpty(sheetName))
            {
                CreateSheet(sheetName);
            }
            lastLineNumber = 0;
        }

        /// <summary>
        /// Create a new sheet in the currently opened Workbook.
        /// </summary>
        /// <param name="name">name of the sheet</param>
        public void CreateSheet(string name)
        {
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild(new Sheets());

            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = name
            };
            sheets.Append(sheet);
            sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            spreadsheetDocument.Save();
        }

        /// <summary>
        /// Write the line content in the file.
        /// </summary>
        /// <param name="lineNumber">number of the line to write</param>
        /// <param name="lineContent">content of the line</param>
        public void WriteLine(int lineNumber, string[] lineContent, bool save = true)
        {
            for (int i = 0; i < lineContent.Count(); i++)
            {
                // @FIXME Improve the performance by directly creating / getting an existing Row object and getting/adding cells inside
                Cell cell = InsertCellInWorksheet(GetExcelColumnName(i + 1), (uint)lineNumber, worksheetPart);

                cell.CellValue = new CellValue(RemoveHexadecimalSymbols(lineContent[i]));
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
            }
            if (save)
            {
                Save();
            }
        }

        /// <summary>
        /// Add a new row to the file.
        /// </summary>
        /// <param name="lineContent">to insert in the row</param>
        public void AppendLine(string[] lineContent)
        {
            lastLineNumber++;
            Row row = new Row() { RowIndex = (uint)lastLineNumber };

            for (int i = 0; i < lineContent.Length; i++)
            {
                string cellReference = GetExcelColumnName(i + 1) + lastLineNumber;
                Cell newCell = new Cell
                {
                    CellReference = cellReference,
                    CellValue = new CellValue(RemoveHexadecimalSymbols(lineContent[i])),
                    DataType = new EnumValue<CellValues>(CellValues.String)
                };
                row.AppendChild(newCell);
            }

            sheetData.Append(row);
        }

        /// <summary>
        /// Save the Excel file.
        /// </summary>
        public void Save()
        {
            worksheetPart.Worksheet.Save();
        }

        /// <summary>
        /// Create a new Excel file and insert the values in a worksheet.
        /// </summary>
        /// <param name="docName"></param>
        /// <param name="allValues"></param>
        public static void CreateNewFileWithValues(string fileName, string worksheetName, List<string[]> allValues)
        {
            using ExcelManager manager = new ExcelManager(fileName, worksheetName);
            for (int j = 0; j < allValues.Count; j++)
            {
                manager.AppendLine(allValues[j]);
            }
            manager.Save();
        }

        /// <summary>
        /// Close the file.
        /// </summary>
        private void Close()
        {
            spreadsheetDocument.Close();
        }

        /// <summary>
        /// Insert a cell in the worksheet.
        /// If it already exists, return it.
        /// </summary>
        /// <param name="columnName"></param>
        /// <param name="rowIndex"></param>
        /// <param name="worksheetPart"></param>
        /// <returns>The cell matching {columnName:rowIndex}</returns>
        /// <remarks>Function found on https://docs.microsoft.com/en-us/office/open-xml/how-to-insert-text-into-a-cell-in-a-spreadsheet?source=recommendations and edited a bit to check parameters.</remarks>
        private Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            if (string.IsNullOrEmpty(columnName))
            {
                throw new ArgumentNullException(nameof(columnName) + " cannot be empty.");
            }
            if (rowIndex == 0)
            {
                throw new ArgumentNullException(nameof(rowIndex) + " must be bigger than 0.");
            }

            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        /// <summary>
        /// Get the Excel column name from the column number.
        /// Example: 1 returns A, 2 returns B etc.
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <returns></returns>
        /// <remarks>Function found on https://stackoverflow.com/a/182924</remarks>
        private string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnName;
        }

        /// <summary>
        /// Remove all the invalid characters in XML to avoid errors.
        /// </summary>
        /// <param name="text"></param>
        /// <returns>the text without invalid character</returns>
        /// <remarks>Function found on https://stackoverflow.com/a/21053139</remarks>
        private static string RemoveHexadecimalSymbols(string text)
        {
            string r = "[\x00-\x08\x0B\x0C\x0E-\x1F\x26]";
            return Regex.Replace(text, r, "", RegexOptions.Compiled);
        }
    }
}
