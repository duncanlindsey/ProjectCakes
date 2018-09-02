# cakes-excel-automation
An add-in for Excel to automate repetitive/cumbersome tasks

## Installation
1. Download "Cakes by Duncan Lindsey v#.#.xlam"
2. Save the file to C:\Users\[your username]\AppData\Roaming\Microsoft\AddIns\ - note: the "AppData" is a hidden folder by default and may not be visible. To find out how to view hidden folders, go to https://support.microsoft.com/en-gb/help/14201/windows-show-hidden-files
3. Open Excel and navigate to File > Options > Add-ins, select "Manage: Excel Add-ins" and click "Go".
4. Check the checkbox next to "Cakes by Duncan Lindsey v#.#" and click "OK".

## How to use
The add-in provides all of its functionality through shortcuts. These shortcuts can be turned on and off and shouldn't override any standard Excel shortcut. 
Turn on shortcuts with **Ctrl + Q**, and off with **Ctrl + Shift + Q**.
Once on, append the following commands to **Ctrl + Shift...** to run the desired process.

## New in version 3.0
Updated look and feel to Cover, Divider and Content worksheets.

**Ctrl + Shift + ...** | **Macro**     | **Description**
-----------------------|---------------|---------------------------
HOME | Jump To Home | Jumps to the first sheet in the workbook
END | Jump To End | Jumps to the last sheet in the workbook
2 | Format Number EUR | Formats cell(s) to EUR formatting
3 | Format Number GBP | Formats cell(s) to GBP formatting
4 | Format Number USD | Formats cell(s) to USD formatting
5 | Format Number Percentage | Formats cell(s) to % with 1 d.p. and bracketed negatives
6 | Format Number Integer | Formats cell(s) to numbers with 0 d.p. and commas for every 000
7 | Add / setup worksheet | Choose from adding an everyday worksheet, to a divider sheet, to a cover sheet with hyperlink-enabled contents page. All sheets come pre-formatted with title, Arial font at size 9 and centred vertical alignment.
8 | Add Subtitle | Creates a subtitle at the selected row, with the option to include a list of sources
9 | Add Subsubtitle | Creates a subsubtitle at the selected row
c | Update Contents | Updates the Table of Contents on the "Cover" sheet, if you have made one already using "Add / setup worksheet"
; (semi-colon) | Font Colour Toggle (Simple) | Toggles font colour between plain black, blue, red and green
' (apostrophe) | Font Colour Toggle (Themes) | Toggles font colour between theme colours, as specified by the design setup of your workbook
/ | Fill Colour Toggle (Dark Themes) | Toggles cell fill colour between dark theme colours, as specified by the design of your workbook (will also change font colour to white)
\# | Fill Colour Toggle (Light Themes) | Toggles cell fill colour between light versions of the theme colours, as specified by the design of your workbook
e | Error Trap | Wraps the formula contained with the selected cell(s) in IF(ISNA(..., IF(ISERR(... or IFERROR(, depending on whether you would like to filter for #NA errors, all errors except #NA errors, or all errors including #NA errors. Allows you to specify the value to return if an error is found. If said return value is a string, don't forget to wrap it in double quote marks.
g | Mass Chart Editor | Make the same formatting changes to multiple charts at once (selected charts, all charts in a sheet, or all charts in a workbook)
x | Check Cell Formatter | Apply tick / exclamation conditional formatting to a check cell, showing a tick if cell value equals 0 (exact), or is within a specified range either side of 0 (sensitivity)
r | Format Total Series (Chart) | Quickly convert the last data series of a chart into a total series (halving the y-axis)
t | Create Crosstab Table (Basic) | Creates simple table formatted within the selected range, including header font and colour and underline
y | Create Crosstab Table (Formal) | Creates a more refined table within the selected range, including offset title and units reminder
u | Create Crosstab Table (Financial Year) | Creates a FY## table within the selected range, where the first FY## must be specified after creation
i | Sheet Importer | Launches the worksheet importer application, allowing multiple sheets from multiple workbooks to be imported into a single workbook with a variety of filtering/selection and naming options (e.g. import all worksheets from a given directory)
l | Increase Decimal | Increases the number of decimal places of the selected cell(s)
k | Decrease Decimal | Decreases the number of decimal places of the selected cell(s)
. (full stop) | Increase Font Size |  Increase the font size of the selected cell(s) by 1 pt
, (comma) | Decrease Font Size |  Decrease the font size of the selected cell(s) by 1 pt

## Troubleshooting
1. If your add-in doesn't work on closing and re-opening Excel
    1. Navigate to C:\Users\[your username]\AppData\Roaming\Microsoft\AddIns\
    2. Right click on "Cakes by Duncan Lindsey v#.#" and choose "Properties"
    3. Check the checkbox against "Unblock"
    4. *Further note: sometimes, even if "Cakes by Duncan Lindsey v#.#" appears ticked in the Add-ins list, one is required to "Browse", select "Cakes by Duncan Lindsey v#.#" again, add, and confirm the intention to replace.*
2. I can't undo... These macros are irreversible and you won't be able to undo past a macro-induced operation. TAKE CARE, and save your work often.
3. Bugs... If you get an error message, choose "End". My apologies, I have not taken the time to properly treat all errors yet. It's on the to-do-list. I'll get around to it soon... maybe...