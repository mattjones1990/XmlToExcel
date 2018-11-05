using System;
using System.Collections.Generic;
using System.Text;

namespace XMLtoEXCEL
{
    class StringManipulation
    {
        public static string GetListOfErroredTests(List<string> errorTests)
        {
            string errors = "";
            foreach (var e in errorTests)
            {
                errors += e + ", ";
            }

            if (errors.Length > 0)
                errors = FormatListOfErroredTests(errors);

            return errors;
        }

        public static string FormatListOfErroredTests(string errors)
        {
                return errors.Remove(errors.Length - 2);
        }
    }
}
