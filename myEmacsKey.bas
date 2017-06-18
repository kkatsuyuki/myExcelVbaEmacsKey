Sub MyEmacsMode()
   With Application
      .OnKey "^{v}" ' paste
      .OnKey "^{y}" ' paste
      .OnKey "^{z}" ' undo
      .OnKey "^{x}" ' cut
      .OnKey "^{u}", "ScrollUp"
      .OnKey "^{d}", "ScrollDown"
      .OnKey "^{b}", "BackwardCell"
      .OnKey "^{p}", "PreviousLine"
      .OnKey "^{f}","ForwardCellModified"
      .OnKey "^{n}","NextLineModified"
      .OnKey "^%{f}", "MoveCol"
      .OnKey "^%{b}", "MoveCol"
      .OnKey "^%{n}", "MoveRow"
      .OnKey "^%{p}", "MoveRow"
      .OnKey "%{g}", "MoveRowCol"
      .OnKey "^{l}", "Recenter"
      .OnKey "^{u}", "ScrollUp"
      .OnKey "^{d}", "ScrollDown"
      .OnKey "^%{j}","SmallScrollDown"
      .OnKey "^%{k}","SmallScrollUp"
      .OnKey "^%{h}","SmallScrollLeft"
      .OnKey "^%{l}","SmallScrollRight"
      .OnKey "^{TAB}","SheetForward"
      .OnKey "^+{TAB}","SheetPrevious"
      .OnKey "^{k}","KillRow"
      .OnKey "^+{k}","KillMultipleRow"
      .OnKey "^{i}","InsertRow"
      .OnKey "^+{i}","InsertMultipleRow"
      .OnKey "^{a}", "BeginningOfUsedRangeLine"
      .OnKey "^{e}", "EndOfUsedRangeLine"
      .OnKey "%{<}","BeginningOfUsedRange"
      .OnKey "%{>}","EndOfUsedRange"
      .OnKey "^%{a}","BeginningOfUsedRangeRow"
      .OnKey "^%{e}","EndOfUsedRangeRow"
      .OnKey "^{t}","CreateSheet"
      .OnKey "^{s}", "Search"
      .OnKey "^{r}", "Replace"
      .OnKey "%{s}","MySaveFile"
      .OnKey "^%{s}","MyWriteFile"
      .OnKey "^%{r}","MyFindFile"
      .OnKey "+{ESC}", "Enable_Keys"
   End With
End Sub

' -------------------------------------------------------------------------
' brought from the original EmacsMode.bas
' -------------------------------------------------------------------------
' forward Cell
Sub ForwardCell()
    If ActiveCell.MergeCells Then
        If ActiveCell.MergeArea.End(xlToRight).Column = Columns.Count Then Exit Sub
    End If
    If ActiveCell.Column <> Columns.Count Then ActiveCell.Offset(0, 1).Activate
End Sub
' backward Cell
Sub BackwardCell()
    If ActiveCell.Column <> 1 Then ActiveCell.Offset(0, -1).Activate
End Sub
' move to the upper line
Sub PreviousLine()
    If ActiveCell.Row <> 1 Then ActiveCell.Offset(-1, 0).Activate
End Sub
' move to the next line
Sub NextLine()
    If ActiveCell.MergeCells Then
        If ActiveCell.MergeArea.End(xlDown).Row = Rows.Count Then Exit Sub
    End If
    If ActiveCell.Row <> Rows.Count Then ActiveCell.Offset(1, 0).Activate
End Sub
' move to the first column in the used range
Sub BeginningOfUsedRangeLine()
    Cells(ActiveCell.Row, ActiveSheet.UsedRange.Column).Activate
End Sub
' move to the last column in the used range
Sub EndOfUsedRangeLine()
    Cells(ActiveCell.Row, _
        ActiveSheet.UsedRange.Columns _
        (ActiveSheet.UsedRange.Columns.Count).Column).Activate
End Sub
' move to the first row in the used range
Sub BeginningOfUsedRangeRow()
    Cells(ActiveSheet.UsedRange.Row, ActiveCell.Column).Activate
End Sub
' move to the last column in the used range
Sub EndOfUsedRangeRow()
    Cells(ActiveSheet.UsedRange.Rows _
        (ActiveSheet.UsedRange.Rows.Count).Row, _
        ActiveCell.Column).Activate
End Sub
' Scroll up to one screen
Sub ScrollDown()
   Dim RowNum As Long
   Dim ColNum As Long
   With ActiveWindow
    RowNum = .ActiveCell.Row - .VisibleRange.Row + 1
    ColNum = .ActiveCell.Column
    .LargeScroll down:=1
    .VisibleRange.Cells(RowNum, ColNum).Activate
   End With
End Sub
' Scroll down to one screen
Sub ScrollUp()
   Dim RowNum As Long
   Dim ColNum As Long
   With ActiveWindow
    RowNum = .ActiveCell.Row - .VisibleRange.Row + 1
    ColNum = .ActiveCell.Column
    .LargeScroll up:=1
    .VisibleRange.Cells(RowNum, ColNum).Activate
   End With
End Sub
' Scroll up/down to put the selection to the middle of the rows
Sub Recenter()
    Dim x As Long
    With ActiveWindow
        x = Int(ActiveCell.Row - (.VisibleRange.Height / ActiveCell.Height) / 2.5)
        If x > 0 Then
            .ScrollRow = x
        End If
    End With
End Sub
 ' Open the search dialog
Sub Search()
   Application.Dialogs(xlDialogFormulaFind).Show
End Sub

' Open the replace dialog
Sub Replace()
   Application.Dialogs(xlDialogFormulaReplace).Show
End Sub


' http://www.rondebruin.nl/key.htm
Sub Enable_Keys()
    Dim StartKeyCombination As Variant
    Dim KeysArray As Variant
    Dim Key As Variant
    Dim I As Long

    On Error Resume Next

    'Shift key = "+"  (plus sign)
    'Ctrl key = "^"   (caret)
    'Alt key = "%"    (percent sign
    'We fill the array with this keys and the key combinations
    'Shift-Ctrl, Shift- Alt, Ctrl-Alt, Shift-Ctrl-Alt

    For Each StartKeyCombination In Array("+", "^", "%", "+^", "+%", "^%", "+^%")

        KeysArray = Array("{BS}", "{BREAK}", "{CAPSLOCK}", "{CLEAR}", "{DEL}", _
                    "{DOWN}", "{END}", "{ENTER}", "~", "{ESC}", "{HELP}", "{HOME}", _
                    "{INSERT}", "{LEFT}", "{NUMLOCK}", "{PGDN}", "{PGUP}", _
                    "{RETURN}", "{RIGHT}", "{SCROLLLOCK}", "{TAB}", "{UP}")

        'Enable the StartKeyCombination key(s) with every key in the KeysArray
        For Each Key In KeysArray
            Application.OnKey StartKeyCombination & Key
        Next Key

        'Enable the StartKeyCombination key(s) with every other key
        For I = 0 To 255
            Application.OnKey StartKeyCombination & Chr$(I)
        Next I

        'Enable the F1 - F15 keys in combination with the Shift, Ctrl or Alt key
        For I = 1 To 15
            Application.OnKey StartKeyCombination & "{F" & I & "}"
        Next I

    Next StartKeyCombination


    'Enable the F1 - F15 keys
    For I = 1 To 15
        Application.OnKey "{F" & I & "}"
    Next I

    'Enable the PGDN and PGUP keys
    Application.OnKey "{PGDN}"
    Application.OnKey "{PGUP}"
End Sub
' -------------------------------------------------------------------------

' modified forward cell to make enabled on integrated cell
Sub ForwardCellModified()
    If ActiveCell.MergeCells Then
        If ActiveCell.MergeArea.Column + ActiveCell.MergeArea.Columns.Count -1 = Columns.Count Then Exit Sub
    End If
    If ActiveCell.Column <> Columns.Count Then ActiveCell.Offset(0, 1).Activate
End Sub
' modified next line to make enabled on integrated cell
Sub NextLineModified()
    If ActiveCell.MergeCells Then
        If ActiveCell.MergeArea.Row + ActiveCell.MergeArea.Rows.Count -1 = Rows.Count Then Exit Sub
    End If
    If ActiveCell.Row <> Rows.Count Then ActiveCell.Offset(1, 0).Activate
End Sub

' move to desired position
Sub MoveRowCol()
   Dim targetRange As Range
   On Error GoTo myError
   Set targetRange = Application.InputBox("Move cell to:",Type:=8,Left:=0,Top:=0)
   On Error GoTo 0
   targetRange.Select
myError:
   Err.Clear
End Sub

Sub MoveCol()
   Dim col As String
   Dim row As String
   Dim colrow As String
   On Error GoTo myError                ' catch error
   col = Application.InputBox("Move column to:",Type:=2,Left:=0,Top:=-100)
   If col = "False" Then
      Exit Sub
   End If
   row = CStr(ActiveCell.Row)
   colrow = col & row
   Range(colrow).Select
   On Error GoTo 0                      ' release catching error
myError:
   Err.Clear                            ' clear error
End Sub

Sub MoveRow()
   Dim col As Long
   Dim row As Long
   On Error GoTo myError                ' catch error
   row = Application.InputBox("Move row to:",Type:=1,Left:=0,Top:=-100)
   If row = False Then
      Exit Sub
   End If
   col = ActiveCell.Column
   Cells(row,col).Select
   On Error GoTo 0                      ' release catching error
myError:
   Err.Clear                            ' clear error
End Sub

' smallScroll
Sub SmallScrollRight()
   ActiveWindow.SmallScroll ToRight:=1
End Sub

Sub SmallScrollLeft()
   ActiveWindow.SmallScroll ToLeft:=1
End Sub

Sub SmallScrollUp
   ActiveWindow.SmallScroll Up:=1
End Sub

Sub SmallScrollDown
   ActiveWindow.SmallScroll Down:=1
End Sub

' change Sheet
Sub SheetForward()
   Dim id As Long
   Dim newId As Long
   id = ActiveSheet.Index
   If Worksheets.Count = id Then
      newId = 1
   Else
      newId = id + 1
   End If

   Worksheets(newId).Activate
End Sub

Sub SheetPrevious()
   Dim id As Long
   Dim newId As Long
   id = ActiveSheet.Index
   If id = 1 Then
      newId = Worksheets.Count
   Else
      newId = id -1
   End If

   Worksheets(newId).Activate
End Sub

' kill Rows
Sub KillRow()
   Rows(ActiveCell.Row).Delete
End Sub
      
Sub KillMultipleRow()
   Dim num As Long
   Dim i As Integer
   num = Application.InputBox("Number of killed rows:",Type:=1,Left:=0,Top:=0)
   
   For i = 1 To num
      Rows(ActiveCell.Row).Delete
   Next

End Sub

' insert Rows
Sub InsertRow()
   Rows(ActiveCell.Row).Insert
End Sub
   
Sub InsertMultipleRow()
   Dim num As Long
   Dim i As Integer
   num = Application.InputBox("Number of rows to be inserted:",Type:=1,Left:=0,Top:=0)
   
   For i = 1 To num
      Rows(ActiveCell.Row).Insert
   Next
End Sub

' select the first cell in used range
Sub BeginningOfUsedRange()
    Cells(ActiveSheet.UsedRange.Row, ActiveSheet.UsedRange.Column).Activate
End Sub
' select the last cell in used range
Sub EndOfUsedRange()
    Cells(ActiveSheet.UsedRange.Rows _
        (ActiveSheet.UsedRange.Rows.Count).Row, _
          ActiveSheet.UsedRange.Columns _
          (ActiveSheet.UsedRange.Columns.Count).column).Activate
    
End Sub

' create new sheet after active sheet
Sub CreateSheet()
   Dim NewWorkSheet As Worksheet
   Dim SheetName As String

   SheetName = Application.InputBox("New sheet name",Type:=2)
   Set NewWorkSheet = Worksheets.Add(After:=ActiveSheet)

   On Error Resume Next
   NewWorkSheet.Name = SheetName
   On Error Goto 0
End Sub


' CxMode
' Sub MyCxMode()
'     With Application
'         .OnKey "^{w}", "MyWriteFile"
'         .OnKey "^{g}", "MyEmacsMode"
'         .OnKey "^{e}", "MyEmacsMode"
'         .OnKey "^{f}", "MyFindFile"
'         .OnKey "^{s}", "MySaveFile"
'         .OnKey "^{x}" ' cut
'         .OnKey "^{v}" ' paste
'         .OnKey "^{z}" ' undo
'     End With
' End Sub

Sub MyWriteFile
   Application.Dialogs(xlDialogSaveAs).Show
   MyEmacsMode
End Sub

Sub MyFindFile()
    Application.Dialogs(xlDialogOpen).Show
    MyEmacsMode
End Sub

' save file
Sub MySaveFile()
   ActiveWorkbook.Save
   MyEmacsMode
End Sub

Sub MyPrintFile()
    Application.Dialogs(xlDialogPrint).Show
    MyEmacsMode
End Sub
