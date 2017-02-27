Sub MyEmacsMode()
   EmacsMode
   With Application
      .OnKey "^{v}" ' paste
      .OnKey "^{z}" ' undo
      .OnKey "^{f}","ForwardCellModified"
      .OnKey "^{n}","NextLineModified"
      .OnKey "^%{f}", "MoveCol"
      .OnKey "^%{b}", "MoveCol"
      .OnKey "^%{n}", "MoveRow"
      .OnKey "^%{p}", "MoveRow"
      .OnKey "%{g}", "MoveRowCol"
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
      .OnKey "%{<}","BeginningOfUsedRange"
      .OnKey "%{>}","EndOfUsedRange"
      .OnKey "^%{a}","BeginningOfUsedRangeRow"
      .OnKey "^%{e}","EndOfUsedRangeRow"
      .OnKey "^{t}","CreateSheet"
      .OnKey "^{x}", "MyCxMode"
      .OnKey "%{s}","MySaveFile"
      .OnKey "^%{r}","MyFindFile"
   End With
End Sub

' modified forward cell
Sub ForwardCellModified()
    If ActiveCell.MergeCells Then
        If ActiveCell.MergeArea.Column + ActiveCell.MergeArea.Columns.Count -1 = Columns.Count Then Exit Sub
    End If
    If ActiveCell.Column <> Columns.Count Then ActiveCell.Offset(0, 1).Activate
End Sub

Sub NextLineModified()
    If ActiveCell.MergeCells Then
        If ActiveCell.MergeArea.Row + ActiveCell.MergeArea.Rows.Count -1 = Rows.Count Then Exit Sub
    End If
    If ActiveCell.Row <> Rows.Count Then ActiveCell.Offset(1, 0).Activate
End Sub

' move to desired position
Sub MoveRowCol()
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
Sub MyCxMode()
    With Application
        .OnKey "^{w}", "MyWriteFile"
        .OnKey "^{g}", "MyEmacsMode"
        .OnKey "^{e}", "MyEmacsMode"
        .OnKey "^{f}", "MyFindFile"
        .OnKey "^{s}", "MySaveFile"
        .OnKey "^{x}" ' cut
        .OnKey "^{v}" ' paste
        .OnKey "^{z}" ' undo
    End With
End Sub

Sub MyWriteFile
   WriteFile
   MyEmacsMode
End Sub

Sub MyFindFile()
    Application.Dialogs(xlDialogOpen).Show
    MyEmacsMode
End Sub

' save file
Sub MySaveFile()
   ThisWorkbook.Save
   ' ThisWorkbook.Saved = True
   ' ActiveWorkbook.Saved = True
   MyEmacsMode
End Sub