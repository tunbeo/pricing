using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;


namespace PricingService
{
    public static class Libs
    {
        public static String Number2String(int number, bool isCaps)
        {
            Char c = (Char)((isCaps ? 65 : 97) + (number - 1));
            return c.ToString();
        }
        public static string GetNextRef(string ref_)
        {
            var matchStart = Regex.Match(ref_, @"([a-zA-Z]+)(\d+)");
            var ro = int.Parse(matchStart.Groups[2].Value);

            char nextChar = (char)((int)ref_[0] + 1);
            var newStr = nextChar.ToString();

            var nextRef = newStr.ToString() + ro.ToString();
            return nextRef;
        }

        public static int GetRowIndexFromRef(string ref_)
        {
            var matchStart = Regex.Match(ref_, @"([a-zA-Z]+)(\d+)");
            var ro = int.Parse(matchStart.Groups[2].Value);
            return ro;
        }

        public static string GetColumnNameFromRef(string ref_)
        {
            var matchStart = Regex.Match(ref_, @"([a-zA-Z]+)(\d+)");
            var col = matchStart.Groups[1].Value;
            return col;
        }

        public static void ToPrintConsole(DataTable dataTable)
        {
            int x = 30;
            string y = "| {0,-30}";
            // Print top line
            Console.WriteLine(new string('-', 5 * x));

            // Print col headers
            var colHeaders = dataTable.Columns.Cast<DataColumn>().Select(arg => arg.ColumnName);
            foreach (String s in colHeaders)
            {
                Console.Write(y, s.Replace("\n", ""));
            }
            Console.WriteLine();

            // Print line below col headers
            Console.WriteLine(new string('-', 5 * x));

            // Print rows
            foreach (DataRow row in dataTable.Rows)
            {
                foreach (Object o in row.ItemArray)
                {
                    Console.Write(y, o.ToString());
                }
                Console.WriteLine();
            }

            // Print bottom line
            Console.WriteLine(new string('-', 5 * x));
        }

        public static int NumberFromExcelColumn(string column)
        {
            int retVal = 0;
            string col = column.ToUpper();
            for (int iChar = col.Length - 1; iChar >= 0; iChar--)
            {
                char colPiece = col[iChar];
                int colNum = colPiece - 64;
                retVal = retVal + colNum * (int)Math.Pow(26, col.Length - (iChar + 1));
            }
            return retVal;
        }

        
        public static double ConvertToDouble(string Value)
        {
            if (Value == null)
            {
                return 0;
            }
            else
            {
                double OutVal;
                double.TryParse(Value, out OutVal);

                if (double.IsNaN(OutVal) || double.IsInfinity(OutVal))
                {
                    return 0;
                }
                return OutVal;
            }
        }

        public static string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }
        public static int ColumnLetterToColumnIndex(string columnLetter)
        {
            columnLetter = columnLetter.ToUpper();
            int sum = 0;

            for (int i = 0; i < columnLetter.Length; i++)
            {
                sum *= 26;
                sum += (columnLetter[i] - 'A' + 1);
            }
            return sum;
        }
    }
}
