# cakes-excel-automation
An add-in for Excel to automate repetitive/cumbersome tasks

# Installation
1. Download "Cakes by Duncan Lindsey v#.#.xlam"
2. Save the file to C:\Users\[your username]\AppData\Roaming\Microsoft\AddIns\ - note: the "AppData" is a hidden folder by default and may not be visible. To find out how to view hidden folders, go to https://support.microsoft.com/en-gb/help/14201/windows-show-hidden-files
3. Open Excel and navigate to File > Options > Add-ins, select "Manage: Excel Add-ins" and click "Go".
4. Check the checkbox next to "Cakes by Duncan Lindsey v#.#" and click "OK".

# How to use
The add-in provides all of its functionality through shortcuts. These shortcuts can be turned on and off and shouldn't override any standard Excel shortcut. 
Turn on shortcuts with **Ctrl + Q**, and off with **Ctrl + Shift + Q**.
Once on, append the following commands to **Ctrl + Shift...** to run the desired process.

**Ctrl + Shift + ...** | **Macro** | **Description**
-----------------------|-------------|---------------------------
HOME                   | Jump To Home | Jumps to the first sheet in the workbook
END                    | Jump To End  | Jumps to the last sheet in the workbook
2 "FormatEuro"
3 "FormatGBP"
4 "FormatUSD"
5 "FormatPercentage"
6 "FormatNumber"
7 "SetupWorksheet"
8 "SubTitle"
9 "SubSubTitle"
c "UpdateContents"
; "FontColourToggle"
' "FontColourToggleThemes"
/ "FillColourDarkToggleThemes"
# "FillColourLightToggleThemes"
e "ErrorTrap"
g "MassChartEditor"
x "CheckCells"
r "FormatTotalSeries"
t "ConstructTableBasic"
y "ConstructTableFormal"
u "ConstructTableFinancial"
i "ImportSheets"
l "Increase_Decimal"
k "Decrease_Decimal"
.  "IncreaseFontSize"
, "DecreaseFontSize"

