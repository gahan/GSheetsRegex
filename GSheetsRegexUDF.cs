using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace GSheetsRegex
{
    public class GSheetsRegexUDF : XlCall
    {
        public static object REGEXMATCH(
            [ExcelArgument(Name = "text", Description = "The text to be tested against the regular expression.")] string sSource,
            [ExcelArgument(Name = "regular_expression", Description = "The regular expression to test the text against.")] string sExpression
        )
        {
            bool bResult = false;

            try
            {
                // Create the reference to the output array of cells

                var oCaller = Excel(xlfCaller) as ExcelReference;
                if (oCaller == null)
                {
                    return ExcelError.ExcelErrorRef;
                }

                // Check that the calling reference is a single cell

                if (oCaller.RowFirst != oCaller.RowLast || oCaller.ColumnFirst != oCaller.ColumnLast) { return ExcelError.ExcelErrorRef; }  // Formula has been entered as an array formula

                Regex oRegex = new Regex(sExpression);
                bResult = oRegex.IsMatch(sSource);
            }
            catch
            {
                return ExcelError.ExcelErrorRef;
            }

            return bResult;
        }

        public static object REGEXMATCHCOUNT(
            [ExcelArgument(Name = "text", Description = "The text to be tested against the regular expression.")] string sSource,
            [ExcelArgument(Name = "regular_expression", Description = "The regular expression to test the text against.")] string sExpression
        )
        {
            int iCounter = 0;

            try
            {
                // Create the reference to the output array of cells

                var oCaller = Excel(xlfCaller) as ExcelReference;
                if (oCaller == null)
                {
                    return ExcelError.ExcelErrorRef;
                }

                // Check that the calling reference is a single cell

                if (oCaller.RowFirst != oCaller.RowLast || oCaller.ColumnFirst != oCaller.ColumnLast) { return ExcelError.ExcelErrorRef; }  // Formula has been entered as an array formula

                Regex oRegex = new Regex(sExpression);
                MatchCollection oMatch = oRegex.Matches(sSource);
                iCounter = oMatch.Count;
            }
            catch
            {
                return ExcelError.ExcelErrorRef;
            }

            return iCounter;
        }

    }
}
