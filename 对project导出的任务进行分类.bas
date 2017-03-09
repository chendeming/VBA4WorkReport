Attribute VB_Name = "IM模块1"

'删除空行
Sub DeleteEmptyRows()
Dim LastRow As Long, r As Long
LastRow = ActiveSheet.UsedRange.Rows.Count
LastRow = LastRow + ActiveSheet.UsedRange.Row - 1
For r = LastRow To 1 Step -1
If WorksheetFunction.CountA(Rows(r)) = 0 Then Rows(r).Delete
Next r
End Sub

'删除空列
Sub DeleteEmptyColumns()
Dim LastColumn As Long, c As Long
LastColumn = ActiveSheet.UsedRange.Columns.Count
LastColumn = LastColumn + ActiveSheet.UsedRange.Column
For c = LastColumn To 1 Step -1
If WorksheetFunction.CountA(Columns(c)) = 0 Then Columns(c).Delete
Next c
End Sub

' 任务分类 宏
Sub 任务分类()
    Call DeleteEmptyRows
    Columns("D:F").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[2]=1,RC[3],R[-1]C)"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[1]<R[1]C[1],0,1)"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=SEARCH(TRIM(RC[1]),RC[1])"
    Range("D2:F2").Select
    
    Dim maxcolumns As Integer
    maxcolumns = ActiveSheet.UsedRange.Rows.Count
    Selection.AutoFill Destination:=Range("D2:F" & CStr(maxcolumns)), Type:=xlFillDefault
    

'    ActiveWindow.SmallScroll Down:=-66
    Range("D2").Select
End Sub

