Attribute VB_Name = "ASP_BAL_Printing"
Public Sub ASP_BAL_Printing()

invoiceRE_TO_balPBI

SLAcomp_TO_balPBI

dashboard_ForVolumne

dashboard1_ForSLA

dashboard2_ForSLA

dashboard3_ForSLA

End Sub


Sub invoiceRE_TO_balPBI()
Dim src As Worksheet
Dim dst As Worksheet
Dim lastrow As Long
Dim location As Range
Dim rng As Range
Dim selection1 As Range
Dim selection2 As Range

Workbooks.Open ("C:\Users\darrell.declaro\Documents\Automation ver1.0\ASP - BAL Printing\BAL Printing Combined (PBI).xlsx")
Workbooks.Open ("C:\Users\darrell.declaro\Documents\Automation ver1.0\ASP - BAL Printing\Raw Data\BALPrint_Invoicing_ReferenceEntry.xlsx")

Set src = Workbooks("BALPrint_Invoicing_ReferenceEntry.xlsx").Worksheets(1)
Set dst = Workbooks("BAL Printing Combined (PBI).xlsx").Worksheets(5)
lastrow = Cells(Rows.Count, 1).End(xlUp).row

checking_for_ERRORS

'************
src.Range("A1").AutoFilter field:=1, Criteria1:="KNCEB-SSC-BALPrinting"
src.Activate
src.Range("B2").Select
ActiveCell.Offset(1, 0).Select
Do Until ActiveCell.EntireRow.Hidden = False
    ActiveCell.Offset(1, 0).Select
Loop

tckId = ActiveCell.Value
Set location = dst.Range("B:B").Find(tckId)

'*********
Set rng = src.Range("a1").CurrentRegion
Set rng = rng.Offset(1, 0)
Set rng = rng.Resize(rng.Rows.Count - 1)
rng.Copy
dst.Range(location.Address).Offset(, -1).PasteSpecial xlPasteValues
Application.CutCopyMode = False
dst.Activate

lastrow = dst.Cells(dst.Rows.Count, 1).End(xlUp).row
Set selection1 = dst.Range("AS2:BD2")
Set selection2 = dst.Range("AS2:BD" & lastrow)

selection1.autofill Destination:=selection2


Workbooks("BALPrint_Invoicing_ReferenceEntry.xlsx").Close savechanges:=False

End Sub



Sub SLAcomp_TO_balPBI()
Dim src As Worksheet
Dim dst As Worksheet
Dim srclastrow As Long
Dim dstlastrow As Long
Dim location As Range
Dim rng As Range
Dim selection1 As Range
Dim selection2 As Range

Workbooks.Open ("C:\Users\darrell.declaro\Documents\Automation ver1.0\ASP - BAL Printing\Raw Data\ASP BAL Printing - For SLA Computation.xlsx")

Set src = Workbooks("ASP BAL Printing - For SLA Computation.xlsx").Worksheets(1)
Set dst = Workbooks("BAL Printing Combined (PBI).xlsx").Worksheets(6)
Set tckId = src.Range("A3")
Set location = dst.Range("A:A").Find(tckId)

srclastrow = src.Cells(src.Rows.Count, 1).End(xlUp).row

src.Range("A3:H" & srclastrow).Copy
dst.Range(location.Address).PasteSpecial xlPasteValues
Application.CutCopyMode = False
dst.Activate

lastrow = dst.Cells(dst.Rows.Count, 1).End(xlUp).row
Set selection1 = dst.Range("I2:W2")
Set selection2 = dst.Range("I2:W" & lastrow)

selection1.autofill Destination:=selection2


Workbooks("ASP BAL Printing - For SLA Computation.xlsx").Close savechanges:=False

End Sub


Sub dashboard_ForVolumne()

Dim src1 As Worksheet
Dim dst1 As Worksheet
Dim dst2 As Worksheet
Dim dst3 As Worksheet
Dim srclastrow As Long


Set src = Workbooks("BAL Printing Combined (PBI).xlsx").Worksheets(4)
Set dst1 = Workbooks("BAL Printing Combined (PBI).xlsx").Worksheets(1)
Set dst2 = Workbooks("BAL Printing Combined (PBI).xlsx").Worksheets(2)
Set dst3 = Workbooks("BAL Printing Combined (PBI).xlsx").Worksheets(3)

ActiveWorkbook.RefreshAll



'working on dashboard  For Volume
srclastrow = src.Cells(src.Rows.Count, 1).End(xlUp).row

dst1.Range("B2:T1000000").ClearContents
dst2.Range("B2:T1000000").ClearContents
dst3.Range("B2:T1000000").ClearContents


src.Range("A3:Q" & srclastrow).Copy
dst1.Range("B2").PasteSpecial xlPasteValues
dst2.Range("B2").PasteSpecial xlPasteValues
dst3.Range("B2").PasteSpecial xlPasteValues
Application.CutCopyMode = False


dst1.Range("A2:A" & srclastrow - 1) = "For Volume"
dst2.Range("A2:A" & srclastrow - 1) = "For Volume"
dst3.Range("A2:A" & srclastrow - 1) = "For Volume"


End Sub


Sub dashboard1_ForSLA()

Dim src As Worksheet
Dim dst1 As Worksheet
Dim srclastrow As Long
Dim dstlastrow1 As Long
Dim dstlastrow1a As Long
Dim selection1 As Range
Dim selection2 As Range


Set src = Workbooks("BAL Printing Combined (PBI).xlsx").Worksheets(4)
Set dst1 = Workbooks("BAL Printing Combined (PBI).xlsx").Worksheets(1)


ActiveWorkbook.RefreshAll


'working on dashboard 1  For SLA
srclastrow = src.Cells(src.Rows.Count, 20).End(xlUp).row
dstlastrow1 = dst1.Cells(dst1.Rows.Count, 2).End(xlUp).Offset(1).row

src.Range("T4:AG" & srclastrow).Copy
dst1.Range("B" & dstlastrow1).PasteSpecial xlPasteValues
Application.CutCopyMode = False

src.Range("AH4:AI" & srclastrow).Copy
dst1.Range("S" & dstlastrow1).PasteSpecial xlPasteValues
Application.CutCopyMode = False


dstlastrow1a = dst1.Cells(dst1.Rows.Count, 2).End(xlUp).row

colA1 = Range("A" & dstlastrow1).Address
colA2 = Range("A" & dstlastrow1a).Address
dst1.Range(colA1 & ":" & colA2) = "For SLA"

dstlastrow1b = dst1.Cells(dst1.Rows.Count, 1).End(xlUp).row
Set selection1 = dst1.Range("U2:V2")
Set selection2 = dst1.Range("U2:V" & dstlastrow1b)
selection1.autofill Destination:=selection2

End Sub




Sub dashboard2_ForSLA()

Dim src As Worksheet
Dim dst1 As Worksheet
Dim srclastrow As Long
Dim dstlastrow1 As Long
Dim dstlastrow1a As Long
Dim selection1 As Range
Dim selection2 As Range


Set src = Workbooks("BAL Printing Combined (PBI).xlsx").Worksheets(4)
Set dst1 = Workbooks("BAL Printing Combined (PBI).xlsx").Worksheets(2)

ActiveWorkbook.RefreshAll


'working on dashboard 2  For SLA
srclastrow = src.Cells(src.Rows.Count, 20).End(xlUp).row
dstlastrow1 = dst1.Cells(dst1.Rows.Count, 2).End(xlUp).Offset(1).row

src.Range("AM4:AZ" & srclastrow).Copy
dst1.Range("B" & dstlastrow1).PasteSpecial xlPasteValues
Application.CutCopyMode = False

src.Range("BA4:BB" & srclastrow).Copy
dst1.Range("S" & dstlastrow1).PasteSpecial xlPasteValues
Application.CutCopyMode = False


dstlastrow1a = dst1.Cells(dst1.Rows.Count, 2).End(xlUp).row

colA1 = Range("A" & dstlastrow1).Address
colA2 = Range("A" & dstlastrow1a).Address
dst1.Range(colA1 & ":" & colA2) = "For SLA"

dstlastrow1b = dst1.Cells(dst1.Rows.Count, 1).End(xlUp).row
Set selection1 = dst1.Range("U2:V2")
Set selection2 = dst1.Range("U2:V" & dstlastrow1b)
selection1.autofill Destination:=selection2

End Sub





Sub dashboard3_ForSLA()

Dim src As Worksheet
Dim dst1 As Worksheet
Dim srclastrow As Long
Dim dstlastrow1 As Long
Dim dstlastrow1a As Long
Dim selection1 As Range
Dim selection2 As Range


Set src = Workbooks("BAL Printing Combined (PBI).xlsx").Worksheets(4)
Set dst1 = Workbooks("BAL Printing Combined (PBI).xlsx").Worksheets(3)

ActiveWorkbook.RefreshAll


'working on dashboard 3  For SLA
srclastrow = src.Cells(src.Rows.Count, 20).End(xlUp).row
dstlastrow1 = dst1.Cells(dst1.Rows.Count, 2).End(xlUp).Offset(1).row

src.Range("BF4:BS" & srclastrow).Copy
dst1.Range("B" & dstlastrow1).PasteSpecial xlPasteValues
Application.CutCopyMode = False

src.Range("BT4:BU" & srclastrow).Copy
dst1.Range("S" & dstlastrow1).PasteSpecial xlPasteValues
Application.CutCopyMode = False


dstlastrow1a = dst1.Cells(dst1.Rows.Count, 2).End(xlUp).row

colA1 = Range("A" & dstlastrow1).Address
colA2 = Range("A" & dstlastrow1a).Address
dst1.Range(colA1 & ":" & colA2) = "For SLA"

dstlastrow1b = dst1.Cells(dst1.Rows.Count, 1).End(xlUp).row
Set selection1 = dst1.Range("U2:V2")
Set selection2 = dst1.Range("U2:V" & dstlastrow1b)
selection1.autofill Destination:=selection2

Workbooks("BAL Printing Combined (PBI).xlsx").Close savechanges:=True

End Sub




Sub checking_for_ERRORS()

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

'assigning 1 to incorrect values
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


selection2.Copy
wsSrc.Range("AR2").PasteSpecial xlPasteValues

selection2.ClearContents

End Sub














