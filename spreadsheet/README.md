# Spreadsheet Pro (Java Swing)

Excel-style spreadsheet in pure Java (JDK 8+, no external libraries).

## Compile and run

Linux / macOS:
    chmod +x build.sh && ./build.sh        # compiles and starts the app

Windows:
    build.bat

Manual:
    mkdir out
    javac -encoding UTF-8 -d out *.java
    java -cp out SpreadsheetApp

## Architecture

| File                        | Layer                                                       |
|-----------------------------|-------------------------------------------------------------|
| SpreadsheetApp.java         | UI: ribbon, formula bar, tabs, status bar, CSV import/export |
| SpreadsheetModel.java       | Data model (AbstractTableModel): raw / formula / style maps  |
| FormulaEngine.java          | Recursive-descent formula evaluator + error codes            |
| SpreadsheetCellRenderer.java| Cell rendering (bold, italic, alignment, colours, borders)   |
| CellStyle.java              | Per-cell formatting record                                   |
| InsertDialog.java           | Insert date / random data / sales table / chart placeholder  |
| MacroIDEDialog.java         | VBA editor + console                                         |
| MacroEngine.java            | VBA-subset interpreter                                       |
| WorkbookIO.java             | File layer: save / load .spx workbooks, optional encryption  |
| SheetProtection.java        | Sheet protection state + salted PBKDF2 password hash         |
| SheetImage.java             | A picture on a sheet (PNG bytes + position + size)           |
| ImageLayer.java             | Draws pictures over the grid; select / drag / resize         |

## Formulas
Start with `=`. Operators `+ - * / ^ %` and parentheses; references `A1`; ranges `A1:A10`.
Functions: SUM, AVERAGE, MIN, MAX, COUNT (ranges and lists), plus ABS, SQRT, ROUND.
Errors: #DIV/0!, #SYNTAX!, #NAME?, #ERROR! (bad reference, text used as number, circular reference).

## Macro example
    Sub Demo()
        For i = 1 To 10
            Cells(i, 1).Value = i * i
            If i Mod 2 = 0 Then Cells(i, 1).Interior.Color = RGB(198, 229, 209)
        Next i
        Range("A11").Formula = "=SUM(A1:A10)"
        Range("A11").Font.Bold = True
    End Sub
Supported: Dim, For/Next (Step, Exit For), If/ElseIf/Else/End If (and single-line If),
Cells(r,c) / Range("A1:B3") with .Value .Formula .Font.Bold/Italic/Size/Color .Interior.Color
.HorizontalAlignment, ClearContents, Clear, MsgBox, Debug.Print, and Int Abs Sqr Rnd Round
Len UCase LCase Str Val Mid RGB.

## Shortcuts
Ctrl+B bold, Ctrl+I italic, Ctrl+L/E/R align, Delete clear, F2 / double-click edit,
Ctrl+N new sheet. Double-click a sheet tab to rename it; the x closes it.

## Files (File menu)
- New Workbook (Ctrl+N), Open... (Ctrl+O), Save (Ctrl+S), Save As... (Ctrl+Shift+S).
- Workbooks are saved as `.spx` and keep all sheets, values, formulas, formatting, pictures
  and sheet protection. The title bar shows `*` for unsaved changes and you are asked
  before anything is discarded.
- Set Open Password...: encrypts the saved file (AES-256-GCM, key from PBKDF2). The password is
  asked when the file is opened. A forgotten password cannot be recovered.
- Import CSV / Export CSV still work as before.

## Pictures
- Insert > Picture from File... (or the Picture button): PNG, JPG, GIF, BMP.
- Insert > Picture from Clipboard: pastes an image copied from another program.
- Click a picture to select it, drag to move, drag the small square in its corner to resize
  (hold Shift to stretch freely), press Delete to remove it.

## Sheet protection (Review menu or the Protect Sheet button)
- Protect with a password, or leave it blank to protect without one.
- A protected sheet cannot be edited, formatted, cleared, auto-summed, filled by a macro or
  have its pictures changed or deleted, and cannot be deleted. A padlock shows on its tab.
- Unprotect Sheet asks for the password. Passwords are never stored, only a salted hash.
- Sheet protection stops changes inside the app (like Excel). To keep a file's contents
  private use Set Open Password, which encrypts it.
