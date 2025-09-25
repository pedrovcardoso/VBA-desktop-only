# VBA-desktop-only

Allows editing Excel workbooks only in the desktop application, while keeping the workbook protected when opened in Excel Online.  
The VBA code locks/unlocks cells and intercepts keyboard shortcuts.

### How it works
- All worksheets are automatically protected when the workbook opens.
- Only the active cell can be edited.
- Intercepts:
  - Ctrl+C → controlled copy
  - Ctrl+V → controlled paste
  - Tab and Shift+Tab → customized navigation
- Works only in Excel Desktop. In Excel Online the workbook remains locked.

### Installation
1. Open the VBA Editor in Excel (Alt+F11).
2. Import the event code file `ThisWorkbook.cls` into the VBA project.
3. Import the helper module `Module1.bas` into the VBA project.
4. Change the default password in the constant `yourPassword` in Module1.
5. Save the file as `.xlsm`.
6. Close and reopen the workbook to activate the code.
7. When opening the workbook, a security warning may appear: "Macros have been disabled" (yellow bar at the top of Excel).  
   Click "Enable Content" so the VBA code works properly.


### Limitations
- Copy & paste will only transfer **text values**, not formatting or styles.
- Moving a cell or range (drag & drop) is not allowed.
- It is blocked only for editing in Excel Online.  
  - Users may still change formatting (colors, styles, etc.), but this behavior can be adjusted in the code.

### Notes
- The password is stored in the VBA code. Do not use sensitive passwords.
- May impact performance on very large workbooks.
- Compatible with Excel 2007 or later, both 32-bit and 64-bit.