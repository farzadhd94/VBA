Attribute VB_Name = "Module1"
Sub AWV_deletingRows()

Dim ws As Worksheet


Set ws = ActiveWorkbook.ActiveSheet

Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim i As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
For i = LastRow To 2 Step -1
    If ws.Cells(i, "H") Like "* SCI*" Then
        ws.Rows(i).Delete
    ElseIf ws.Cells(i, "H") Like "* U.A.*" Then
        ws.Rows(i).Delete
    ElseIf ws.Cells(i, "H") Like "*DE*" Then
        ws.Rows(i).Delete
    ElseIf ws.Cells(i, "H") Like "* Sarl*" Then
        ws.Rows(i).Delete
        ElseIf ws.Cells(i, "H") Like "* SA*" Then
        ws.Rows(i).Delete
        ElseIf ws.Cells(i, "H") Like "*01_*" Then
        ws.Rows(i).Delete
        ElseIf ws.Cells(i, "H") Like "* SLU*" Then
        ws.Rows(i).Delete
        ElseIf ws.Cells(i, "H") Like "*Wustermarker*" Then
        ws.Rows(i).Delete
    End If

Next i
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub
