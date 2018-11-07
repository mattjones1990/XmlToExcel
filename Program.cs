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
                //Create Worksheets
                var excelWorksheet1 = ExcelFactory.CreateWorksheet("High Level Report", excel);
                var excelWorksheet2 = ExcelFactory.CreateWorksheet("Development View Report", excel);

                //Add headers
                ExcelFactory.GenerateHeaders(excelWorksheet1, excelWorksheet2);

                int worksheet1Row = 2;
                int worksheet2Row = 2;
                int worksheet2TestRow = 2;

                foreach (XmlNode node in testsuiteList)
                {
                    int addInitialRowOfData = 0;
                    //First Worksheet - High Level
                    #region
                    //Console stats 
                    totalTests = totalTests + XmlFactory.GetStatsForConsole(node, "tests");
                    totalFailures = totalFailures + XmlFactory.GetStatsForConsole(node, "failures");

                    //Add testsuite name
                    excelWorksheet1.Cells["A" + worksheet1Row.ToString()].Value = testSuitesName;

                    //Add scenario definition
                    excelWorksheet1.Cells["B" + worksheet1Row.ToString()].Value = XmlFactory.GetInnerText(node, "name");

                    //Add test numbers
                    totalTests = XmlFactory.GetInnerTextInt(node, "tests");
                    totalTestsConsole = totalTestsConsole + totalTests;
                    excelWorksheet1.Cells["C" + worksheet1Row.ToString()].Value = totalTests;

                    //Add test status
                    excelWorksheet1.Cells["D" + worksheet1Row.ToString()].Value = XmlFactory.PassOrFailCheck(node, "failures");

                    //Add failed test descriptions
                    List<string> errorTests = XmlFactory.GetErroredTests(node);
                    excelWorksheet1.Cells["E" + worksheet1Row.ToString()].Value = StringManipulation.GetListOfErroredTests(errorTests);

                    worksheet1Row++;
                    #endregion

                    //Second Worksheet
                    #region

                    XmlNodeList testCaseList = node.SelectNodes("testcase"); //working
                   
                    foreach (XmlNode item in testCaseList)
                    {
                        XmlNodeList failedTestCases = item.SelectNodes("failure");
                        bool failCount = failedTestCases.Count > 0;

                        if (failCount && addInitialRowOfData == 0)
                        {
                            //Add testsuite name
                            excelWorksheet2.Cells["A" + worksheet2Row.ToString()].Value = testSuitesName;

                            //Add scenario definition
                            excelWorksheet2.Cells["B" + worksheet2Row.ToString()].Value = XmlFactory.GetInnerText(node, "name");
                            addInitialRowOfData++;
                        }

                        if (failCount)
                        {
                            //Add test names and associated errors/stacktraces
                            excelWorksheet2.Cells["C" + worksheet2Row.ToString()].Value = XmlFactory.GetInnerText(item, "name");
                            excelWorksheet2.Cells["D" + worksheet2Row.ToString()].Value = StringManipulation.SortErrors(failedTestCases[0].InnerXml, "Error message");
                            excelWorksheet2.Cells["E" + worksheet2Row.ToString()].Value = StringManipulation.SortErrors(failedTestCases[0].InnerXml, "Stacktrace");

                            worksheet2Row++;
                        }
                    }
                    #endregion
                }

                //Save Excel file
                ExcelFactory.SaveSpreadsheet(directory, excel);
            }

            WritingTextOutput.TestStats(totalTestsConsole, totalFailures);
            Console.ReadLine();
        }
    }
}
