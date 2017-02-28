# myExcelVbaEmacsKey

My configuration of Excel vba to use Emacs key bindings

## Features

-   Based on EmacsMode.bas

-   Added some useful function
    
    -   `SheetForward` / `SheetPrevious` : Select the next/previous sheet
    -   `KillRow` : Kill the active row
    -   `MoveRow` / `MoveCol` : Move the selection toward the row/column direction by inputted number/alphabet
    
    See [Default key bindings and function descriptions (including original commands written in `EmacsMode.bas`)](#org63247c8) about other functions.
-   Remaped the key bindings
    
    See [Default key bindings and function descriptions (including original commands written in `EmacsMode.bas`)](#org63247c8) in details.

## Install

### Use only on the specified book

Import the module files (`EmacsMode.bas`, `myEmacsKey.bas`) as the standard module and execute the procedure `MyEmacsMode`.

1.  Open the Excel book.
2.  Click Developer>"Visual Basic" on Ribbon (or Alt + F11).
    
    → VBE (Visual Basic Editor) is opened
3.  Select the desired project in Project explorer and File>"Import File&#x2026;" on menu bar (or right click on the desired project and click "Import File&#x2026;")
4.  Select the module files (`EmacsMode.bas`, `myEmacsKey.bas`) and import.
    
    → New modules are created under the modules hierarchy on "Project explorer"
5.  Return to the Excel book window and click Developer>macro on Ribbon and execute `MyEmacsMode`.

### 

## Default key bindings and function descriptions (including original commands written in `EmacsMode.bas`)

| Key Binding | Function Name           | Description                                                        |
|----------- |----------------------- |------------------------------------------------------------------ |
| `C-f`       | ForwardCellModified     | select forward cell                                                |
| `C-b`       | BackwardCell            | select backward cell                                               |
| `C-n`       | NextLineModified        | select next line                                                   |
| `C-n`       | PreviousLine            | select previous line                                               |
| `C-M-f`     | MoveCol                 | Move the selection to the row directed by the inputted number      |
| `C-M-b`     | MoveCol                 | Move the selection to the row directed by the inputted number      |
| `C-M-n`     | MoveRow                 | Move the selection to the column directed by the inputted alphabet |
| `C-M-p`     | MoveRow                 | Move the selection to the column directed by the inputted alphabet |
| `M-g`       | MoveRowCol              | Move the selection to the cell directed by inputted ID             |
| `C-l`       | Recenter                | Scroll up/down to center the active row                            |
| `C-M-j`     | SmallScrollDown         | Scroll down by one row                                             |
| `C-M-k`     | SmallScrollUp           | Scroll up by one row                                               |
| `C-M-l`     | SmallScrollRight        | Scroll right by one column                                         |
| `C-M-h`     | SmallScrollLeft         | Scroll left by one column                                          |
| `C-TAB`     | SheetForward            | Select the forward sheet                                           |
| `C-S-TAB`   | SheetPrevious           | Select the previous sheet                                          |
| `C-k`       | KillRow                 | Kill the active row                                                |
| `C-S-k`     | KillMultipleRow         | Kill the multiple rows by the inputted number                      |
| `C-i`       | InsertRow               | Insert a new row                                                   |
| `C-S-i`     | InsertMultipleRow       | Insert the multiple rows by the inputted number                    |
| `M-<`       | BeginningOfUsedRange    | Select the first cell in the used range                            |
| `M->`       | EndOfUsedRange          | Select the last cell in the used range                             |
| `C-M-a`     | BeginningOfUsedRangeRow | Move the selection to the first row in the used range              |
| `C-M-e`     | EndOfUsedRangeRow       | Move the selection to the last row in the used range               |
| `C-t`       | CreateSheet             | Create the new sheet you named after the active sheet              |
| `C-s`       | Search                  | Open the search dialog                                             |
| `M-s`       | MySaveFile              | Save the book                                                      |
| `C-M-r`     | MyFindFile              | Open the dialog and select the file to be opened                   |
| `C-x`       | MyCxMode                | The command to change the keymap to use the command starting `C-x` |
| `C-x C-s`   | MySaveFile              | Save the book                                                      |
| `C-x C-f`   | MyFindFile              | Open the dialog and select the file to be opened                   |
| `C-x C-w`   | MyWriteFile             | Save the book as another name                                      |
| `C-x C-g`   | MyEmacsMode             | Activate this emacs key bindings                                   |
| `C-x C-e`   | MyEmacsMode             | Activate this emacs key bindings                                   |
| `S-ESC`     | Enable\_Keys            | Deactivate this emacs key bindings                                 |

## Modify

Since the configuration meets only my needs, I encourage you to modify some configurations especially about key bindings. Modifying is enabled only by modifying imported module directly on VBE or by importing the module file (`myEmacsKey.bas`) you edited.
