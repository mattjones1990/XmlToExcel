using System;
using System.Collections.Generic;
using System.Text;

namespace XMLtoEXCEL
{
    class WritingTextOutput
    {
        public static void StartText(string dir)
        {
            Console.WriteLine("---- XML to Excel tool ----");
            Console.WriteLine("---------------------------");
            Console.WriteLine("---------------------------");
            Console.WriteLine("---------------------------\n");
            Console.WriteLine("XML file must be in " + dir);
            Console.WriteLine("Once you have one XML file in this directory, press enter.");
        }

        public static void TestStats(int totaltests, int totalfailures)
        {
            Console.WriteLine("\nTotal Tests: " + totaltests.ToString());
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Total Passes: " + (totaltests - totalfailures).ToString());
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Total Failures: " + (totalfailures).ToString());
            Console.ForegroundColor = ConsoleColor.Gray;
        }
    }
}
