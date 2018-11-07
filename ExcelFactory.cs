using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using OfficeOpenXml;

namespace XMLtoEXCEL
{
    class ExcelFactory
    {
        /// <summary>
        /// Add worksheet object to the method below and copy the headerRow & headerRange code to suit the new worksheet.
        /// </summary>
        /// <param name="excelWorksheet1"></param>
        /// <param name="excelWorksheet2"></param>
        public static void GenerateHeaders(ExcelWorksheet excelWorksheet1, ExcelWorksheet excelWorksheet2)
        {
            List<string[]> headerRow1 = new List<string[]>()
                {
                    new string[] { "Testsuites Name", "Scenario Definition", "Tests", "Test Status", "Failed Test Description" }
                };

            List<string[]> headerRow2 = new List<string[]>()
                {
                    new string[] { "Testsuites Name", "Scenario Definition", "Failed Test Decription", "Error Message", "StackTrace"
                        //"Error Messages", "testScriptError", "Stack Track Message", "View Report" 
                    }
                };

            string headerRange1 = "A1:" + Char.ConvertFromUtf32(headerRow1[0].Length + 64) + "1";
            string headerRange2 = "A1:" + Char.ConvertFromUtf32(headerRow2[0].Length + 64) + "1";

            excelWorksheet1.Cells[headerRange1].LoadFromArrays(headerRow1);
            excelWorksheet2.Cells[headerRange2].LoadFromArrays(headerRow2);

            excelWorksheet1.Cells[headerRange1].Style.Font.Bold = true;
            excelWorksheet2.Cells[headerRange2].Style.Font.Bold = true;
        }
        public static ExcelWorksheet CreateWorksheet(string name, ExcelPackage spreadsheet)
        {
            return spreadsheet.Workbook.Worksheets.Add(name);
        }
        public static void SaveSpreadsheet(string directory, ExcelPackage excel)
        {
            var date = DateTime.Now;
            FileInfo excelFile = new FileInfo(directory + "/" + date.Hour + date.Minute + "_" + date.Day +  "_" + date.Month + "_" + date.Year + "_PostmanTestOutput.xlsx");
            excel.SaveAs(excelFile);
        }
    }
}
