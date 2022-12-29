Attribute VB_Name = "test_convert_toNumber"
Sub convert_toNumber()
Dim selection1 As Range
Dim selection2 As Range
Dim lastrow As Long
Dim cell As Range
Dim colm As Range

lastrow = Cells(Rows.Count, 1).End(xlUp).row

'search and remove for EB's
ActiveSheet.Range("B:B").Replace What:="EB's", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False


'convert cell value to number
Range("C2").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-1]*1)"

Set selection1 = Range("C2")
Set selection2 = Range("C2:C" & lastrow)

selection1.autofill Destination:=selection2

selection2.Copy
selection2.PasteSpecial xlPasteValues

'assigning 1 to incorrect values
Set colm = Application.Selection

For Each cell In colm
    If cell.Value > 10000 And cell.Value <> "" Then
        cell.Value = 1
    End If
Next cell

Range("C2:C" & lastrow).Copy
Range("B2").PasteSpecial xlPasteValues

Range("C2:C" & lastrow).ClearContents

End Sub

