Attribute VB_Name = "Module2"
Sub awv_ColumnsOrder()
Dim search As Range
Dim cnt As Integer
Dim colOrdr As Variant
Dim indx As Integer


Dim sheetname As String
sheetname = InputBox("When the data is downloaded?")
ActiveSheet.Name = sheetname

ActiveSheet.Copy After:=Worksheets(Sheets.Count)
sheetname = InputBox("What is your sheet's name?")
ActiveSheet.Name = sheetname


colOrdr = Array("Status", "Kontobezeichnung", "Kontoinhaber", "Buchungsdatum", "Betrag", "Währung", "FiBu-Kontonummer", "Buchungskreis", "Verwendungszweck", "Partner Name", "IBAN") 'define column order with header names here

cnt = 1


For indx = LBound(colOrdr) To UBound(colOrdr)
    Set search = Rows("1:1").Find(colOrdr(indx), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    If Not search Is Nothing Then
        If search.Column <> cnt Then
            search.EntireColumn.Cut
            Columns(cnt).Insert Shift:=xlToRight
            Application.CutCopyMode = False
        End If
    cnt = cnt + 1
    End If
Next indx

Columns(cnt - 1).EntireColumn.Insert
Cells(1, cnt - 1).Value = "Land"

Columns(cnt).EntireColumn.Insert
Cells(1, cnt).Value = "Kennzahl G-Vorfall"

Columns(cnt + 1).EntireColumn.Insert
Cells(1, cnt + 1).Value = "Betrag G-Vorfall"

Columns(cnt + 2).EntireColumn.Insert
Cells(1, cnt + 2).Value = "Kennzahl Steuer"

Columns(cnt + 3).EntireColumn.Insert
Cells(1, cnt + 3).Value = "Betrag Steuer"

Columns(cnt + 4).EntireColumn.Insert
Cells(1, cnt + 4).Value = "Bemerkung"

Columns(cnt + 5).EntireColumn.Insert
Cells(1, cnt + 5).Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
Range(Columns(cnt + 7), Columns(cnt + 37)).Delete

End Sub
