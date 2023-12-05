Attribute VB_Name = "Module3"
Sub AWV_nichtmeldepflichtig()

Dim ws As Worksheet


Set ws = ActiveWorkbook.ActiveSheet

Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim i As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
For i = LastRow To 2 Step -1
    If ws.Cells(i, "I") Like "*Liqui*" Then
        ws.Cells(i, 16).Value = "nicht meldepflichtig, kontoübertrag"
        ws.Cells(i, 1).Interior.Color = RGB(248, 203, 173)
    ElseIf ws.Cells(i, "B") Like "*Betriebs*" Then
        ws.Cells(i, 16).Value = "nicht meldepflichtig, Konto in Luxemburg zur Zahlung der nuf BK"
        ws.Cells(i, 1).Interior.Color = RGB(248, 203, 173)
        

    End If

Next i
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

