Step 1: Open the VBA Editor
1. Open your Excel workbook that contains (or will contain) the CMS numeric export data.
2. Press Alt + F11 to open the Visual Basic for Applications (VBA) editor.

Step 2: Insert a New Module
1. In the VBA editor, go to the menu bar → click Insert → Module.
2. A new blank module will appear (usually named Module1).

Step 3: Paste the Code
1. Copy the entire script you posted.
2. Paste it into the new module window.

3. If you want local time correction, edit the line near the top:
    Const TZ_OFFSET_MINUTES As Long = 0
         Change it to:
    Const TZ_OFFSET_MINUTES = -240 for Eastern Daylight Time (EDT)
    Const TZ_OFFSET_MINUTES = -300 for Eastern Standard Time (EST)
    (or any offset in minutes from UTC)

5. Save your VBA project:
- Press Ctrl + S
- If prompted, save as a Macro-Enabled Workbook (.xlsm).

Step 4: Prepare Your Data Sheet
1. Open the sheet that contains your raw CMS numeric export data.
2. Make sure it’s the active sheet — the macro will modify this sheet directly.
3. Optionally, make a backup copy of the sheet first (recommended).

Step 5: Run the Macro
1. Press Alt + F8 to open the Macro List.
2. Select FixCMS_NumericExport from the list.
3. Click Run.

Step 6: Verify Results
- After a few seconds, the macro will:
- Add headers if missing
- Convert UNIX epoch timestamps into readable local times
- Format durations (Talk, Hold, ACW, etc.) into h:mm:ss
- Convert flags (0/1 → Y/N)
- Auto-fit columns
- Turn your data into a nicely formatted Excel table named CMS_Export

Step 6: Verify Results
- After a few seconds, the macro will:
- Add headers if missing
- Convert UNIX epoch timestamps into readable local times
- Format durations (Talk, Hold, ACW, etc.) into h:mm:ss
- Convert flags (0/1 → Y/N)
- Auto-fit columns
- Turn your data into a nicely formatted Excel table named CMS_Export

- Optional: Add a Button (for one-click use)
1. In Excel, go to Developer → Insert → Button (Form Control).
  - If you don’t see the Developer tab, enable it in:
    File → Options → Customize Ribbon → check “Developer”.
2. Draw the button on your sheet.
3. When prompted, assign the macro FixCMS_NumericExport.
4. Label it “Fix CMS Export” (or whatever you like).
5. Now you can just click that button next time you import CMS numeric data.
  

