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
        public static ExcelWorksheet CreateWorksheet(string name, ExcelPackage spreadsheet)
        {
            return spreadsheet.Workbook.Worksheets.Add(name);
        }
        public static void SaveSpreadsheet(string directory, ExcelPackage excel)
        {
            FileInfo excelFile = new FileInfo(directory + "/test.xlsx");
            excel.SaveAs(excelFile);
        }
    }
}
