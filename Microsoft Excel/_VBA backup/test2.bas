Attribute VB_Name = "test2"
Sub test2()
Dim wsSrc As Worksheet
Dim selection1 As Range
Dim selection2 As Range
Dim lastrow As Long
Dim cell As Range
Dim colm As Range

Set wsSrc = Workbooks("BALPrint_Invoicing_ReferenceEntry.xlsx").Worksheets(1)

lastrow = wsSrc.Cells(Rows.Count, "AQ").End(xlUp).row


With wsSrc.Range("AQ2:AQ" & lastrow)
    .NumberFormat = "General"
    .Value = .Value
End With

'search and remove for EB's
wsSrc.Range("AR:AR").Replace What:="EB's", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False



'convert cell value to number
wsSrc.Range("AS2").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-1]*1)"

Set selection1 = wsSrc.Range("AS2")
Set selection2 = wsSrc.Range("AS2:AS" & lastrow)

selection1.autofill Destination:=selection2

selection2.Copy
selection2.PasteSpecial xlPasteValues

For Each cell In wsSrc.Range("AS2:AS" & lastrow)
    If IsError(cell) Then
        cell.Value = 1
    ElseIf IsEmpty(cell) Then
        cell.Value = ""
    ElseIf IsNumeric(cell.Value) Then
        cell.Value = cell.Value
    End If
Next


For Each cell In wsSrc.Range("AS2:AS" & lastrow)
    If cell.Value > 1000 Then
        cell.Value = 1
    End If
Next

End Sub
