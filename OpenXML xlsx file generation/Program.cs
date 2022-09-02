using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXML_xlsx_file_generation
{
    internal class Program
    {
        /// <summary>
        /// Create a new Excel file and insert some values inside.
        /// </summary>
        /// <remarks>The official documentation of DocumentFormat.OpenXml is available at https://docs.microsoft.com/en-us/office/open-xml/how-do-i</remarks>
        static void Main()
        {
            // Declare some data to insert in the file
            List<string[]> rows = new List<string[]>
            {
                new string[] { "Id", "Name" },
                new string[] { "1", "John" },
                new string[] { "2", "Dupond" }
            };
            // Information regarding the file to create
            string fileName = "DataFile.xlsx";
            string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName);
            string sheetName = "Data";
            // Call the function to create the file and insert the values inside
            CreateNewFileWithValues(filePath, sheetName,  rows);
        }


        /// <summary>
        /// Create a new Excel file and insert the values in a worksheet.
        /// </summary>
        /// <param name="docName"></param>
        /// <param name="allValues"></param>
        public static void CreateNewFileWithValues(string fileName, string worksheetName, List<string[]> allValues)
        {
            // Create a new file
            using SpreadsheetDocument spreadSheet = SpreadsheetDocument.
                Create(fileName, SpreadsheetDocumentType.Workbook);
            {
                // Add a Workbook to the document
                WorkbookPart workbookpart = spreadSheet.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add Sheets to the Workbook.
                Sheets sheets = spreadSheet.WorkbookPart.Workbook.
                    AppendChild(new Sheets());

                // Append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet()
                {
                    Id = spreadSheet.WorkbookPart.
                    GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = worksheetName
                };
                sheets.Append(sheet);

                for (int j = 0; j < allValues.Count; j++)
                {
                    var rowText = allValues[j];
                    for (int i = 0; i < rowText.Count(); i++)
                    {
                        // Insert cell A1 into the new worksheet.
                        Cell cell = InsertCellInWorksheet(GetExcelColumnName(i + 1), (uint) (j + 1), worksheetPart);

                        // Set the value of cell A1.
                        cell.CellValue = new CellValue(rowText[i]);
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    }
                }
                // Save the new worksheet.
                worksheetPart.Worksheet.Save();
            }
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
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
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
        private static string GetExcelColumnName(int columnNumber)
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

    }
}
