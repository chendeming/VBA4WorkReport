Attribute VB_Name = "IMģ��1"



'ɾ������
Sub DeleteEmptyRows()
Dim LastRow As Long, r As Long
LastRow = ActiveSheet.UsedRange.Rows.Count
LastRow = LastRow + ActiveSheet.UsedRange.Row - 1
For r = LastRow To 1 Step -1
    Dim countRow As Integer
    countRow = WorksheetFunction.CountA(Rows(r))
    If countRow = 0 Then
     Rows(r).Delete
    End If
Next r
End Sub

'ɾ������
Sub DeleteEmptyColumns()
Dim LastColumn As Long, c As Long
LastColumn = ActiveSheet.UsedRange.Columns.Count
LastColumn = LastColumn + ActiveSheet.UsedRange.Column
For c = LastColumn To 1 Step -1
If WorksheetFunction.CountA(Columns(c)) = 0 Then Columns(c).Delete
Next c
End Sub

' ������� ��
Sub �������()
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
    
    Dim maxRows As Integer
    maxRows = ActiveSheet.UsedRange.Rows.Count
    Selection.AutoFill Destination:=Range("D2:F" & CStr(maxRows)), Type:=xlFillDefault
    

'    ActiveWindow.SmallScroll Down:=-66
    Range("D2").Select
End Sub

Sub ȡ���ϲ���Ԫ��()
'
' ȡ���ϲ���Ԫ�� ��
'
'
    Cells.Select
    Selection.UnMerge
    Range("A1").Select
End Sub


' ERP����ϲ� ��
Sub ERP����ϲ�()

    ' ȡ���ϲ���Ԫ��
    Call ȡ���ϲ���Ԫ��
    
    ' ������
    ActiveSheet.Columns(14).Insert
    
    Cells(1, 14) = "������ϸ"
    
    Dim nowRow As Integer
    nowRow = 2
    Dim strContent As String
    strContent = Cells(nowRow, 13)

    Dim iRow As Integer
    For iRow = nowRow + 1 To ActiveSheet.UsedRange.Rows.Count Step 1
    If Cells(iRow, 1) = "" Then
        strContent = strContent & Chr(10) & Cells(iRow, 13)
        Range("M" & iRow) = Clear
    Else
        Cells(nowRow, 14) = strContent
        nowRow = iRow
        strContent = Cells(nowRow, 13)
    End If
    Next iRow
    
    Cells(nowRow, 14) = strContent

End Sub
