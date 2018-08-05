# GSheetsRegex
A port of the RegEx functions from Google Sheets for Excel using ExcelDNA

This is a small C#/ExcelDNA project to create an Excel Add-In to give the Regex functionality that exists in Google Sheets.  The three main Regex functions: REGEXMAATCH, REGEXEXTRACT & REGEXREPLACE have been ported and one new function created REGEXMATCHCOUNT which returns the number of matchs that are found in a specified text by an expression.  I have also extended the original three functions with some additional, optional parameters giving some additional control over the behaviour of the Regex parsing.  The library that makes the Excel add-in possible is ExcelDNA which can be found here: https://excel-dna.net/

## Installation

Whilst I will be looking to create a proper installer for the add-in for the moment installation will need to be done manually by copying the appropiate (32 or b4 bit depending on your installation of Office) .xll from the bin/release directory to a permanent home on your PC and then following this procedure to install the add-in to your installation of Excel.

* Open Excel Options
* Select the Add-Ins property page
* Select "Manage Excel Add-ins"
* Browse to where you have stored the .xll file and select it

If installed correctly you should then have the four functions available in Excel:
![Regex functions available in Excel](https://github.com/gahan/GSheetsRegex/blob/master/images/Functions%20Available.PNG "Regex functions available in Excel")

## Function Definitions

### REGEXMATCH

#### Parameters

**text** - _string_ - The text to be tested against the regular expression.

**regular_expression** - _string_ - The regular expression to test the text against.

**ignore_case** - _boolean_ - Specifies case-insensitive matching.  TRUE by default.

**multiline** - _boolean_ - Multiline mode. Changes the meaning of ^ and $ so they match at the beginning and end, respectively, of any line, and not just the beginning and end of the entire string.

**right_to_left** - _boolean_ - Specifies that the search will be from right to left instead of from left to right.

**singleline** - _boolean_ - Specifies single-line mode. Changes the meaning of the dot (.) so it matches every character (instead of every character except \n).

#### Returns

_boolean_ - TRUE if the expression can be matched within the text.


### REGEXMATCHCOUNT

#### Parameters

**text** - _string_ - The text to be tested against the regular expression.

**regular_expression** - _string_ - The regular expression to test the text against.

**ignore_case** - _boolean_ - Specifies case-insensitive matching.  TRUE by default.

**multiline** - _boolean_ - Multiline mode. Changes the meaning of ^ and $ so they match at the beginning and end, respectively, of any line, and not just the beginning and end of the entire string.

**right_to_left** - _boolean_ - Specifies that the search will be from right to left instead of from left to right.

**singleline** - _boolean_ - Specifies single-line mode. Changes the meaning of the dot (.) so it matches every character (instead of every character except \n).

#### Returns

_integer_ - The number of times the expression can be matched within the text.


### REGEXEXTRACT

#### Parameters

**text** - _string_ - The text to be tested against the regular expression.

**regular_expression** - _string_ - The regular expression to test the text against.

**match_item** - _integer_ - Optional parameter to select which matching group is returned. 

**ignore_case** - _boolean_ - Specifies case-insensitive matching.  TRUE by default.

**multiline** - _boolean_ - Multiline mode. Changes the meaning of ^ and $ so they match at the beginning and end, respectively, of any line, and not just the beginning and end of the entire string.

**right_to_left** - _boolean_ - Specifies that the search will be from right to left instead of from left to right.

**singleline** - _boolean_ - Specifies single-line mode. Changes the meaning of the dot (.) so it matches every character (instead of every character except \n).

#### Returns

_string_ - The text that matches the expression within the specified string.  By default this will be the first match but this can be changed to whichever match is required by specifying the *match_item* parameter.


### REGEXREPLACE

#### Parameters

**text** - _string_ - The text to be tested against the regular expression.

**regular_expression** - _string_ - The regular expression to test the text against.

**replacement** - _string_ - The text to replace the matched groups within the source text. 

**ignore_case** - _boolean_ - Specifies case-insensitive matching.  TRUE by default.

**multiline** - _boolean_ - Multiline mode. Changes the meaning of ^ and $ so they match at the beginning and end, respectively, of any line, and not just the beginning and end of the entire string.

**right_to_left** - _boolean_ - Specifies that the search will be from right to left instead of from left to right.

**singleline** - _boolean_ - Specifies single-line mode. Changes the meaning of the dot (.) so it matches every character (instead of every character except \n).

#### Returns

_string_ - The text that matches the expression within the specified string.  By default this will be the first match but this can be changed to whichever match is required by specifying the *match_item* parameter.
