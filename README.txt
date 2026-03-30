PREMIUM COMPARISON SYSTEM — README
=====================================
Files: Premium_Comparison_System.vba
       Premium_Worksheet_Events.vba
Version: In active development
Last updated: March 2026

HOW TO INSTALL
--------------
1. Open your Excel workbook
2. Press Alt+F11 to open the VBA editor
3. In the Project panel on the left, find your workbook name
4. Right-click on "Modules" and select "Insert > Module"
5. Open Premium_Comparison_System.vba in any text editor (Notepad works)
6. Select All (Ctrl+A), Copy (Ctrl+C)
7. Click inside the new Module in the VBA editor and Paste (Ctrl+V)
8. Next, find your data sheet in the Project panel
   (it will be listed under your workbook as "Sheet1" or similar)
9. Double-click that sheet to open its code window
10. Open Premium_Worksheet_Events.vba in Notepad
11. Select All (Ctrl+A), Copy (Ctrl+C)
12. Click inside the sheet code window and Paste (Ctrl+V)
13. Close the VBA editor
14. To run Build UI manually: press Alt+F8, select "BuildFullUI", click Run
15. Or click the buttons that appear on your sheet after Build UI runs

REQUIREMENTS
------------
- Excel with macros enabled (go to File > Options > Trust Center >
  Macro Settings and select "Enable all macros" if buttons do not work)
- Both source and target files must have a named header row
  (a row containing column names like ASSETID, LOCATION, DESCRIPTION etc)
- Files with no column header names will not work correctly
- Excel Table formatted files are supported — the macro automatically
  converts them to plain ranges before building the UI

HOW TO USE
----------
1. SETUP
   - Open both your source and target Excel files
   - Open the workbook containing this macro

2. SELECT FILES
   - Click "Select Files" button
   - A numbered list of open workbooks and sheets will appear
   - Enter the number corresponding to your source workbook, then sheet
   - Enter the number corresponding to your target workbook, then sheet
   - Example: if the list shows "1. QA_ASSET  2. PRD_ASSET"
     enter 1 to select QA_ASSET, then enter the sheet number
   - Repeat for target file

3. BUILD UI (first time setup only)
   - Build UI only needs to be run once per sheet to set up the match area
   - If the UI is already built on your sheet, skip this step
   - To build: press Alt+F8, select "BuildFullUI" and click Run
     OR click "Build UI" if the button is already visible on your sheet
   - A prompt will show candidate header rows with a preview of their content
     — enter the row number that contains your actual column names
   - Also enter the header row number for your target file when prompted
   - Confirm the row insertion dialog
   - The UI will be built above your data with match rule rows
   - AutoFilter dropdowns will appear on your data header row automatically
   - Note: if a merge warning popup appears during Build UI, always click
     CANCEL — see Warnings section below

4. CONFIGURE MATCH RULES
   - Each row in the match area is one match rule
   - Double-click a column header cell in the data area to assign it to a
     match rule row (marks with X in red)
   - Double-click again to clear the X
   - Match Type is auto-populated as Match_1, Match_2 etc — rename as needed
   - Use "+ Add Match" to add more match rule rows
   - Use "- Delete" to remove rows (enter the numbers shown in Match column)
   - Use "Clear X" to clear X marks from the currently selected match row

5. EXECUTE MATCH
   - Click "Execute Match" to run the comparison
   - A prompt will appear asking which row is the header row in the source
     file — enter the row number shown in the candidate list
   - A second prompt will ask which row is the header row in the target file
   - A third prompt will ask which column in the TARGET file contains the
     unique ID to write to MATCHED_ID (e.g. ASSETID, ASSETNUM, ID)
     Enter the exact column name as it appears in the target file header
   - Results are written to MATCH_STATUS, MATCHED_ID etc columns
   - DONE = matched, NO_MATCH = not found
   - Large files may take several minutes depending on row count

BUTTONS
-------
+ Add Match   — Add a new match rule row
- Delete      — Delete match rule rows by entering their Match numbers
Clear X       — Clear X marks from the selected match row
Select Files  — Choose source and target files to compare
Check Status  — Show current source and target file connection status
Execute Match — Run the comparison using configured match rules
Build UI      — Rebuild the match UI (first time setup or after changes)
Pause Macro   — Pause/Resume macro events for safe column deletion and Undo

MATCH RESULT COLUMNS
--------------------
MATCHED_ID    — The ID value from the target file that matched
MATCH_TYPE    — Which match rule produced the result
MATCH_STATUS  — DONE, NO_MATCH, or error status
SOURCE_FILE   — Source file name
TARGET_FILE   — Target file name

IMPORTANT WARNINGS
------------------
1. DELETING COLUMNS
   Before deleting any column, click the "Pause Macro" button first.
   This temporarily disables macro events so that:
   - Column deletion will not trigger macro interference
   - Ctrl+Z (Undo) will work normally after deletion
   After you are done deleting columns, click "Pause Macro" again
   (it will say "Resume Macro") to re-enable macro events.
   If you forget to pause and delete a column accidentally,
   close the file WITHOUT saving and reopen from your last manual save.
   Safe practice:
   - Always turn AutoSave OFF before working with this macro
   - Save your file manually before making any column changes

2. MERGE WARNING POPUP DURING BUILD UI
   A popup saying "Merging cells only keeps the upper-left value"
   may appear during Build UI. Always click CANCEL on this popup —
   clicking OK will merge the Match Type cell and delete its content.
   The UI builds correctly regardless when you click Cancel.

3. DELETING MATCH ROWS
   Always use the "- Delete" button to remove match rule rows.
   Manually deleting rows directly in Excel may cause layout issues.
   If this happens accidentally, click "Build UI" to rebuild in place.

4. EXCEL TABLES
   If your file uses Excel Table formatting, the macro automatically
   converts it to a plain range before building the UI.
   Your data is preserved — only the table formatting is removed.

5. FILES WITHOUT NAMED HEADERS
   The macro requires a row of named column headers.
   Files with only data rows and no header row will not work correctly.

KNOWN LIMITATIONS
-----------------
- Execute Match on very large files (50,000+ rows) may be slow.
  This is a known performance area being improved.
- Deleting a column cannot be undone — see Warning 1 above.