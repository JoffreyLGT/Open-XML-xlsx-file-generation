using System;
using System.Collections.Generic;
using System.IO;

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
            string fileName = "DataFile-Automatic.xlsx";
            string folder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string sheetName = "Data";

            // Use the static method to create and write content in the file
            ExcelManager.CreateNewFileWithValues(Path.Combine(folder, fileName), sheetName, rows);

            // Create a new instance of ExcelManager and write specific lines
            fileName = "DataFile-Manual.xlsx";
            using ExcelManager excel = new ExcelManager(folder, fileName, sheetName);
            // Write content only in line 1, 3 and 5
            excel.WriteLine(1, rows[0]);
            excel.WriteLine(3, rows[1]);
            excel.WriteLine(5, rows[2]);

            fileName = "DataFile-AppendLine.xlsx";
            using ExcelManager excelAppendLine = new ExcelManager(folder, fileName, sheetName);
            excelAppendLine.AppendLine(rows[0]);
            for (int i = 0; i < 100000; i++)
            {
                excelAppendLine.AppendLine(new string[] { i.ToString(), $"Line {i}", "Hello there", "how are you today?", "I am fine thank you" });
            }
        }
    }
}
