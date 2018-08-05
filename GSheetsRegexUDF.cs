using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace GSheetsRegex
{
    public class GSheetsRegexUDF : XlCall, IExcelAddIn
    {
        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }

        public void AutoOpen()
        {
            IntelliSenseServer.Install();
        }

        private enum RegexAction
        {
            Match,
            MatchCount,
            Extract,
            Replace
        }

        // Default values for the Options parameters

        private const bool _IGNORECASE = true;
        private const bool _MULTILINE = false;
        private const bool _RIGHTTOLEFT = false;
        private const bool _SINGLELINE = false;

        // User defined function declarations

        public static object REGEXMATCH(
            [ExcelArgument(Name = "text", Description = "The text to be tested against the regular expression.")] string sSource,
            [ExcelArgument(Name = "regular_expression", Description = "The regular expression to test the text against.")] string sExpression,
            [ExcelArgument(Name = "ignore_case", Description = "Specifies case-insensitive matching.  TRUE by default.")] [Optional] [DefaultParameterValue(_IGNORECASE)] bool bIgnoreCase,
            [ExcelArgument(Name = "multiline", Description = "Multiline mode. Changes the meaning of ^ and $ so they match at the beginning and end, respectively, of any line, and not just the beginning and end of the entire string.")] [Optional] [DefaultParameterValue(_MULTILINE)] bool bMultiLine,
            [ExcelArgument(Name = "right_to_left", Description = "Specifies that the search will be from right to left instead of from left to right.")] [Optional] [DefaultParameterValue(_RIGHTTOLEFT)] bool bRightToLeft,
            [ExcelArgument(Name = "singleline", Description = "Specifies single-line mode. Changes the meaning of the dot (.) so it matches every character (instead of every character except \n).")] [Optional] [DefaultParameterValue(_SINGLELINE)] bool bSingleLine
        )
        {
            return GoogleSheetsRegex(RegexAction.Match, sSource, sExpression, null, 0, bIgnoreCase, bMultiLine, bRightToLeft, bSingleLine);
        }

        public static object REGEXMATCHCOUNT(
            [ExcelArgument(Name = "text", Description = "The text to be tested against the regular expression.")] string sSource,
            [ExcelArgument(Name = "regular_expression", Description = "The regular expression to test the text against.")] string sExpression,
            [ExcelArgument(Name = "ignore_case", Description = "Specifies case-insensitive matching.  TRUE by default.")] [Optional] [DefaultParameterValue(_IGNORECASE)] bool bIgnoreCase,
            [ExcelArgument(Name = "multiline", Description = "Multiline mode. Changes the meaning of ^ and $ so they match at the beginning and end, respectively, of any line, and not just the beginning and end of the entire string.")] [Optional] [DefaultParameterValue(_MULTILINE)] bool bMultiLine,
            [ExcelArgument(Name = "right_to_left", Description = "Specifies that the search will be from right to left instead of from left to right.")] [Optional] [DefaultParameterValue(_RIGHTTOLEFT)] bool bRightToLeft,
            [ExcelArgument(Name = "singleline", Description = "Specifies single-line mode. Changes the meaning of the dot (.) so it matches every character (instead of every character except \n).")] [Optional] [DefaultParameterValue(_SINGLELINE)] bool bSingleLine
        )
        {
            return GoogleSheetsRegex(RegexAction.MatchCount, sSource, sExpression, null, 0, bIgnoreCase, bMultiLine, bRightToLeft, bSingleLine);
        }

        public static object REGEXEXTRACT(
            [ExcelArgument(Name = "text", Description = "The text to be tested against the regular expression.")] string sSource,
            [ExcelArgument(Name = "regular_expression", Description = "The regular expression to test the text against.")] string sExpression,
            [ExcelArgument(Name = "match_item", Description = "The nth match of the expression in the source text.")] [Optional] int iMatchNumber,
            [ExcelArgument(Name = "ignore_case", Description = "Specifies case-insensitive matching.  TRUE by default.")] [Optional] [DefaultParameterValue(_IGNORECASE)] bool bIgnoreCase,
            [ExcelArgument(Name = "multiline", Description = "Multiline mode. Changes the meaning of ^ and $ so they match at the beginning and end, respectively, of any line, and not just the beginning and end of the entire string.")] [Optional] [DefaultParameterValue(_MULTILINE)] bool bMultiLine,
            [ExcelArgument(Name = "right_to_left", Description = "Specifies that the search will be from right to left instead of from left to right.")] [Optional] [DefaultParameterValue(_RIGHTTOLEFT)] bool bRightToLeft,
            [ExcelArgument(Name = "singleline", Description = "Specifies single-line mode. Changes the meaning of the dot (.) so it matches every character (instead of every character except \n).")] [Optional] [DefaultParameterValue(_SINGLELINE)] bool bSingleLine
        )
        {
            return GoogleSheetsRegex(RegexAction.Extract, sSource, sExpression, null, (iMatchNumber <= 0 ? 1 : iMatchNumber), bIgnoreCase, bMultiLine, bRightToLeft, bSingleLine);
        }

        public static object REGEXREPLACE(
            [ExcelArgument(Name = "text", Description = "The text to be tested against the regular expression.")] string sSource,
            [ExcelArgument(Name = "regular_expression", Description = "The regular expression to test the text against.")] string sExpression,
            [ExcelArgument(Name = "replacement", Description = "The replacement text which will be inserted into the original text.")] string sReplacement,
            [ExcelArgument(Name = "ignore_case", Description = "Specifies case-insensitive matching.  TRUE by default.")] [Optional] [DefaultParameterValue(_IGNORECASE)] bool bIgnoreCase,
            [ExcelArgument(Name = "multiline", Description = "Multiline mode. Changes the meaning of ^ and $ so they match at the beginning and end, respectively, of any line, and not just the beginning and end of the entire string.")] [Optional] [DefaultParameterValue(_MULTILINE)] bool bMultiLine,
            [ExcelArgument(Name = "right_to_left", Description = "Specifies that the search will be from right to left instead of from left to right.")] [Optional] [DefaultParameterValue(_RIGHTTOLEFT)] bool bRightToLeft,
            [ExcelArgument(Name = "singleline", Description = "Specifies single-line mode. Changes the meaning of the dot (.) so it matches every character (instead of every character except \n).")] [Optional] [DefaultParameterValue(_SINGLELINE)] bool bSingleLine
        )
        {
            return GoogleSheetsRegex(RegexAction.Replace, sSource, sExpression, sReplacement, 0, bIgnoreCase, bMultiLine, bRightToLeft, bSingleLine);
        }

        // Private function that handles all of the processing as much is common throughout each function

        private static object GoogleSheetsRegex(RegexAction iAction, string sSource, string sExpression, string sReplacement, int iMatchNumber, bool bIgnoreCase, bool bMultiLine, bool bRightToLeft, bool bSingleLine)
        {
            try
            {
                if (string.IsNullOrEmpty(sExpression)) { return ExcelError.ExcelErrorNull; }

                // Create the reference to the output array of cells

                var oCaller = Excel(xlfCaller) as ExcelReference;
                if (oCaller == null)
                {
                    return ExcelError.ExcelErrorRef;
                }

                // Check that the calling reference is a single cell

                if (oCaller.RowFirst != oCaller.RowLast || oCaller.ColumnFirst != oCaller.ColumnLast) { return ExcelError.ExcelErrorRef; }  // Formula has been entered as an array formula

                // Assemble the Regex options

                var oRegExOptions = RegexOptions.None;

                if (bIgnoreCase) { oRegExOptions = (oRegExOptions == RegexOptions.None ? RegexOptions.IgnoreCase : oRegExOptions | RegexOptions.IgnoreCase); }
                if (bMultiLine) { oRegExOptions = (oRegExOptions == RegexOptions.None ? RegexOptions.Multiline : oRegExOptions | RegexOptions.Multiline); }
                if (bRightToLeft) { oRegExOptions = (oRegExOptions == RegexOptions.None ? RegexOptions.RightToLeft : oRegExOptions | RegexOptions.RightToLeft); }
                if (bSingleLine) { oRegExOptions = (oRegExOptions == RegexOptions.None ? RegexOptions.Singleline : oRegExOptions | RegexOptions.Singleline); }

                // Perform the Regex matching

                Regex oRegex = new Regex(sExpression, oRegExOptions);

                // Return the appropiate result

                switch (iAction)
                {
                    case RegexAction.Match:
                        return oRegex.IsMatch(sSource);

                    case RegexAction.MatchCount:
                        return oRegex.Matches(sSource).Count;

                    case RegexAction.Extract:
                        MatchCollection oMatches = oRegex.Matches(sSource);
                        return oMatches[(iMatchNumber > oMatches.Count ? oMatches.Count - 1 : iMatchNumber - 1)].Value;

                    case RegexAction.Replace:
                        return oRegex.Replace(sSource, sReplacement);
                }
            }
            catch
            {
                return ExcelError.ExcelErrorRef;
            }

            return null;
        }
    }
}
