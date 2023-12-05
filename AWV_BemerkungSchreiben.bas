Attribute VB_Name = "Module4"
Sub AWV_BemerkungSchreiben()

Dim ws As Worksheet


Set ws = ActiveWorkbook.ActiveSheet

Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim i As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
For i = LastRow To 2 Step -1
    If ws.Cells(i, "J") Like "*SAVILLS COMM*" Then
        ws.Cells(i, 17).Value = "Sweeps fehlen, als Miete gemeldet"
        ws.Cells(i, 17).Interior.Color = RGB(255, 255, 0)
        ws.Cells(i, 1).Interior.Color = RGB(198, 224, 180)
        ws.Cells(i, 11).Value = "IE"
        ws.Cells(i, 12).Value = "280(3)"
        ws.Cells(i, 13).Value = Abs(Cells(i, 5))
        ws.Cells(i, 14).Value = "'---"
        ws.Cells(i, 15).Value = "'---"
        ws.Cells(i, 16).Value = "Bruttokaltmiete (StSchl A0)"

        
    ElseIf ws.Cells(i, "B") Like "*Betriebs*" Then
        ws.Cells(i, 16).Value = "nicht meldepflichtig, Konto in Luxemburg zur Zahlung der nuf BK"
        ws.Cells(i, 1).Interior.Color = RGB(248, 203, 173)
        

    End If

Next i
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub


