using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
//Install-Package EPPlus


namespace XMLtoEXCEL
{
    class Program
    {
        static void Main(string[] args)
        {
            string directory = "C:/temp/XMLtoEXCEL";
            FolderFileChecks checks = new FolderFileChecks(directory);

            //Show opening text
            WritingTextOutput.StartText(directory);
            Console.ReadLine();

            //Validate folder and file structure
            bool successfulCheck = checks.CheckFolderandFiles();

            //If there are issues with the directory or file, error out.
            if (!successfulCheck)
            {
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("\nPress any key to exit.");
                Console.ReadLine();
                return;
            }

            Console.WriteLine("Loading File...");

            FileInfo[] xmlFiles = checks.GetDirectoryFiles();

            XmlDocument xml = new XmlDocument();
            xml.Load(xmlFiles[0].FullName);
            XmlNodeList testsuitesList = xml.SelectNodes("/testsuites");
            XmlNodeList testsuiteList = xml.SelectNodes("/testsuites/testsuite");

            int totalTests = 0;
            int totalFailures = 0;
            int totalTestsConsole = 0;
            string testSuitesName = testsuitesList[0].Attributes.GetNamedItem("name").InnerText;


            using (ExcelPackage excel = new ExcelPackage())
            {
                //First worksheet creation
                var excelWorksheet1 = ExcelFactory.CreateWorksheet("High Level Report", excel);

                List<string[]> headerRow1 = new List<string[]>()
                {
                    new string[] { "Testsuites Name", "Scenario Definition", "Tests", "Test Status", "Failed Test Description" }
                };

                string headerRange1 = "A1:" + Char.ConvertFromUtf32(headerRow1[0].Length + 64) + "1";
                excelWorksheet1.Cells[headerRange1].LoadFromArrays(headerRow1);
                excelWorksheet1.Cells[headerRange1].Style.Font.Bold = true;

                //Second worksheet creation
                var excelWorksheet2 = ExcelFactory.CreateWorksheet("Development View Report", excel);

                List<string[]> headerRow2 = new List<string[]>()
                {
                    new string[] { "Testsuites Name", "Scenario Definition", "Failed Test Decription", "Errors", "Failures",
                        "Error Messages", "testScriptError", "Stack Track Message", "View Report" }
                };

                string headerRange2 = "A1:" + Char.ConvertFromUtf32(headerRow2[0].Length + 64) + "1";
                excelWorksheet2.Cells[headerRange2].LoadFromArrays(headerRow2);
                excelWorksheet2.Cells[headerRange2].Style.Font.Bold = true;

                int testSuiteRow = 2;
                int scenarioDefinitionRow = 2;
                int testNumberDefinition = 2;
                int testStatus = 2;
                int errorStatus = 2;

                foreach (XmlNode node in testsuiteList)
                {
                    //Console stats 
                    totalTests = totalTests + XmlFactory.GetStatsForConsole(node, "tests");
                    totalFailures = totalFailures + XmlFactory.GetStatsForConsole(node, "failures");

                    //Add testsuite name
                    excelWorksheet1.Cells["A" + testSuiteRow.ToString()].Value = testSuitesName;
                    testSuiteRow++;

                    //Add scenario definition
                    excelWorksheet1.Cells["B" + scenarioDefinitionRow.ToString()].Value = XmlFactory.GetInnerText(node, "name");
                    scenarioDefinitionRow++;

                    //Add test numbers
                    totalTests = XmlFactory.GetInnerTextInt(node, "tests");
                    totalTestsConsole = totalTestsConsole + totalTests;
                    excelWorksheet1.Cells["C" + testNumberDefinition.ToString()].Value = totalTests;
                    testNumberDefinition++;

                    //Add test status
                    excelWorksheet1.Cells["D" + testStatus.ToString()].Value = XmlFactory.PassOrFailCheck(node, "failures");
                    testStatus++;

                    //Add failed test descriptions
                    List<string> errorTests = XmlFactory.GetErroredTests(node);
                    excelWorksheet1.Cells["E" + errorStatus.ToString()].Value = StringManipulation.GetListOfErroredTests(errorTests);
                    errorStatus++;
                }

                //Save Excel file
                ExcelFactory.SaveSpreadsheet(directory, excel);
            }

            WritingTextOutput.TestStats(totalTestsConsole, totalFailures);
            Console.ReadLine();
        }


    }
}
