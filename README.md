# myExcelVbaEmacsKey

My configuration of Excel vba to use Emacs key bindings

## Features

-   Based on EmacsMode.bas

-   Added some useful function
    -   `SheetForward` / `PreviousSheet` : Select the next/previous sheet
    -   `KillRow` : Kill the active row
    -   `MoveRow` / `MoveCol` : Move the selection toward the row/column direction by inputted number/alphabet &#x2026; and so on
-   Remaped the key bindings

## Install

### Use only on the specified book

Inport the module files (`EmacsMode.bas`, `myEmacsKey.bas`) as the standard module and execute the procedure `MyEmacsMode`.

1.  Open the Excel book.
2.  Click Developer>"Visual Basic" on Ribbon (Alt + F11).
    
    → VBE (Visual Basic Editor) is opened
3.  Select the desired project in Project explorer and File>"Import File&#x2026;" on menu bar (Right click on the desired project and click "Import File&#x2026;")
4.  Select the module files (`EmacsMode.bas`, `myEmacsKey.bas`) and import. → Create new modules under Modules hierarchy on Project explorer
5.  Return to the Excel book window and click Developer>macro on Ribbon and execute `MyEmacsMode`.

### 

## Default key bindings and function descriptions

| Key Binding | Function Name       | Description                                     |
|----------- |------------------- |----------------------------------------------- |
| `C-f`       | ForwardCellModified | select forward cell                             |
| `C-n`       | NextLineModified    | select next line                                |
| `C-M-f`     | MoveCol             | Move the selection toward the row direction     |
| `C-M-b`     |                     | by inputted number                              |
| `C-M-n`     | MoveRow             | Move the selection toward the column direction  |
| `C-M-p`     |                     | by inputted alphabet                            |
| `M-g`       | MoveRowCol          | Move the selection to the cell of inputted ID   |
| `C-M-j`     | SmallScrollDown     | Scroll down by one row                          |
| `C-M-k`     | SmallScrollUp       | Scroll up by one row                            |
| `C-M-l`     | SmallScrollRight    | Scroll right by one column                      |
| `C-M-h`     | SmallScrollLeft     | Scroll left by one column                       |
| `C-TAB`     | SheetForward        | Select the forward sheet                        |
| `C-S-TAB`   | SheetPrevious       | Select the previous sheet                       |
| `C-k`       | KillRow             | Kill the active row                             |
| `C-S-k`     | KillMultipleRow     | Kill the multiple rows by the inputted number   |
| `C-i`       | InsertRow           | Insert a new row                                |
| `C-S-i`     | InsertMultipleRow   | Insert the multiple rows by the inputted number |
|             |                     |                                                 |

## Modify

Since the configuration meets only my needs, I encourage you to modify some configurations especially about key bindings. Modifying is enabled only by modifying imported module directly on VBE or editing the module file(`myEmacsKey.bas`) and import it. After modifying don't forget to execute `MyEmacsMode`.
