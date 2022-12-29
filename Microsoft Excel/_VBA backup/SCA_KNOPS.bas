Attribute VB_Name = "SCA_KNOPS"
Public Sub SCA_KNOPS()

to_template

to_PBI

End Sub
Sub to_template()

Dim src As Worksheet
Dim dst As Worksheet
Dim srclastrow As Long
Dim dstlastrow As Long
Dim selection1 As Range
Dim selection2 As Range

Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\SCA - KNOPS\Extracted from BO\NAM KNOPS Report.xlsx"
Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\SCA - KNOPS\SCA KNOPS - TEMPLATE.xlsx"

Set src = Workbooks("NAM KNOPS Report.xlsx").Worksheets(2)
Set dst = Workbooks("SCA KNOPS - TEMPLATE.xlsx").Worksheets(1)

ActiveSheet.ShowAllData

srclastrow = src.Cells(src.Rows.Count, 1).End(xlUp).row
dstlastrow = dst.Cells(dst.Rows.Count, 1).End(xlUp).row


'convert text to number
With src.Range("Q2:Q" & srclastrow)
    .NumberFormat = "General"
    .Value = .Value
End With

With src.Range("A2:A" & srclastrow)
    .NumberFormat = "General"
    .Value = .Value
End With

'checking non numeric errors
For Each cell In src.Range("Q2:Q" & srclastrow)
    If IsNumeric(cell) Then
        cell.Value = cell.Value
    Else
        cell.Value = 1
    End If
Next


'to template
dst.Range("A2:R" & dstlastrow).ClearContents

src.Range("A2:R" & srclastrow).Copy
dst.Range("A2").PasteSpecial xlPasteValues
Application.CutCopyMode = False

lastrow = dst.Cells(dst.Rows.Count, 1).End(xlUp).row

dst.Range("S2:AD2").Copy
dst.Range("S2:AD" & lastrow).PasteSpecial xlPasteFormulas

Workbooks("NAM KNOPS Report.xlsx").Close savechanges:=False

MsgBox "check column S for any error and adjust accordingly"
End Sub


Sub to_PBI()

Dim src As Worksheet
Dim dst As Worksheet
Dim srclastrow As Long
Dim dstlastrow As Long
Dim selection1 As Range
Dim selection2 As Range

Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\SCA - KNOPS\SCA KNOPS (PBI).xlsx"

Set src = Workbooks("SCA KNOPS - TEMPLATE.xlsx").Worksheets(1)
Set dst = Workbooks("SCA KNOPS (PBI).xlsx").Worksheets(1)

srclastrow = src.Cells(src.Rows.Count, 1).End(xlUp).row
dstlastrow = dst.Cells(dst.Rows.Count, 1).End(xlUp).row



'Change/update Column R "SSC SJO" based on the countries
src.Range("A2:AD" & srclastrow).AutoFilter field:=28, Criteria1:=Array("for clarification", "no data"), Operator:=xlFilterValues

src.Range("A2:AD" & srclastrow).AutoFilter field:=2, Criteria1:="*|*", Operator:=xlFilterValues

src.Activate
src.Range("R2").Select
ActiveCell.Offset(1, 0).Select
Do Until ActiveCell.EntireRow.Hidden = False
    ActiveCell.Offset(1, 0).Select
Loop

src.Range(Selection, ActiveCell.SpecialCells(xlLastCell).Offset(, -12)).FormulaR1C1 = "=INDEX(Reference!R2C9:R47C9,MATCH(LEFT(RC[-16],FIND(""|"",RC[-16])-2),Reference!R2C11:R47C11,0))"

ActiveSheet.ShowAllData


'filter for region SCA
src.Range("A2:AD" & srclastrow).AutoFilter field:=28, Criteria1:="SCA", Operator:=xlFilterValues


'to pbi
src.Activate
src.Range("A2").Select
ActiveCell.Offset(1, 0).Select
Do Until ActiveCell.EntireRow.Hidden = False
    ActiveCell.Offset(1, 0).Select
Loop

tckId = ActiveCell.Value
Set location = dst.Range("A:A").Find(tckId)

'*********
Set rng = src.Range("A1").CurrentRegion
Set rng = rng.Offset(1, 0)
Set rng = rng.Resize(rng.Rows.Count - 1)
rng.Copy
dst.Range(location.Address).PasteSpecial xlPasteValues
Application.CutCopyMode = False


Workbooks("SCA KNOPS - TEMPLATE.xlsx").Close savechanges:=True


End Sub

