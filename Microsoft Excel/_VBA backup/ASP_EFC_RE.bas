Attribute VB_Name = "ASP_EFC_RE"
Public Sub ASP_ExportFileCreation_ReferenceEntry()

dataPreparation1

dataPreparation2

End Sub

Sub dataPreparation1()
Dim wsSrc1 As Worksheet
Dim wsDst1 As Worksheet
Dim srclastrow1 As Long
Dim dstlastrow1 As Long
Dim selection1 As Range
Dim selection2 As Range
Dim selection3 As Range
Dim selection4 As Range
Dim selection5 As Range
Dim selection6 As Range
Dim lastrow As Long

Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\ASP - Export File Creation + Reference Entry\Raw Data\Export File Creation.xlsx" 'source1
Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\ASP - Export File Creation + Reference Entry\EXP_FC Volume (PBI).xlsx" 'destination1

'copy from Export File Creation and paste values to EXP_FC Volume (PBI)
Set wsSrc1 = Workbooks("Export File Creation.xlsx").Worksheets("Details")
Set wsDst1 = Workbooks("EXP_FC Volume (PBI).xlsx").Worksheets("EXP_FC")

srclastrow1 = wsSrc1.Cells(wsSrc1.Rows.Count, "A").End(xlUp).row
dstlastrow1 = wsDst1.Cells(wsDst1.Rows.Count, "A").End(xlUp).Offset(1).row


'column A
wsSrc1.Range("A4:A" & srclastrow1).Copy
wsDst1.Range("E" & dstlastrow1).PasteSpecial xlPasteValues
Application.CutCopyMode = False

'column W to Y
wsSrc1.Range("W4:Y" & srclastrow1).Copy
wsDst1.Range("A" & dstlastrow1).PasteSpecial xlPasteValuesAndNumberFormats
Application.CutCopyMode = False


lastrow = wsDst1.Cells(Rows.Count, 1).End(xlUp).row
Range("D1").End(xlDown).Offset(1, 11).Select
Selection = "retain"
ActiveCell.autofill Destination:=Range(ActiveCell.Address & ":O" & lastrow)

'autofill
Set selection1 = wsDst1.Range("D2")
Set selection2 = wsDst1.Range("D2:D" & lastrow)
selection1.autofill Destination:=selection2

'autofill
Set selection3 = wsDst1.Range("F2:M2")
Set selection4 = wsDst1.Range("F2:M" & lastrow)
selection3.autofill Destination:=selection4

lastrowA = wsDst1.Cells(wsDst1.Rows.Count, 1).End(xlUp).row
With wsDst1.Range("A2:B" & lastrowA)
    .NumberFormat = "dd/mm/yyyy hh:mm:ss"
    .Value = .Value
End With



Workbooks("Export File Creation.xlsx").Close

removingDuplicatesOpt1

Workbooks("EXP_FC Volume (PBI).xlsx").Close savechanges:=True
End Sub

Sub dataPreparation2()
Dim wsSrc As Worksheet
Dim wsDst As Worksheet
Dim srclastrow As Long
Dim dstlastrow As Long
Dim location As Range

Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\ASP - Export File Creation + Reference Entry\Raw Data\BALPrint_Invoicing_ReferenceEntry.xlsx" 'source2
Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\ASP - Export File Creation + Reference Entry\BAL_INV Volume (PBI).xlsx" 'destination2

'cleaning
checking_for_ERRORS

'locating the cell value from source to destination, then paste
Set wsSrc = Workbooks("BALPrint_Invoicing_ReferenceEntry.xlsx").Worksheets("DOC-E-006")
Set wsDst = Workbooks("BAL_INV Volume (PBI).xlsx").Worksheets("BAL_INV")
Set tckId = wsSrc.Range("B2")
Set location = wsDst.Range("B:B").Find(tckId)

srclastrow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).row

wsSrc.Range("A2:AR" & srclastrow).Copy
wsDst.Range(location.Address).Offset(, -1).PasteSpecial xlPasteValuesAndNumberFormats
Application.CutCopyMode = False

Workbooks("BALPrint_Invoicing_ReferenceEntry.xlsx").Close savechanges:=False


Workbooks("BAL_INV Volume (PBI).xlsx").Close savechanges:=True
End Sub


Sub removingDuplicatesOpt1()

'NOTE:
'remove of duplicates function will start from bottom to top
'any new data appended from the bottom will be remove as duplicates
'thus, sorting data to descending before removing of duplicates
'then sorting to ascending to reutrn data to original form


Dim selection1 As Range
Dim selection2 As Range
Dim lastrow As Long

lastrow = Cells(Cells.Rows.Count, "A").End(xlUp).row



'adding index column
Range("B:B").Insert
Range("B:B").NumberFormat = "General"
Range("B2") = 1
Range("B3") = 2

Set selection1 = ActiveWorkbook.Worksheets(1).Range("B2:B3")
Set selection2 = ActiveWorkbook.Worksheets(1).Range("B2:B" & lastrow)

'autofill
selection1.autofill Destination:=selection2


'sort to descending
ActiveWorkbook.Worksheets(1).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(1).Sort.SortFields.add Key:=Range("B1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(1).Sort
        .SetRange Range("A2:P" & lastrow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
'remove duplicates
ActiveWorkbook.Worksheets(1).Range("A1:P" & lastrow).removeDuplicates Columns:=1, Header:=xlYes


'sort to ascending
ActiveWorkbook.Worksheets(1).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(1).Sort.SortFields.add Key:=Range("B1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(1).Sort
        .SetRange Range("A2:P" & lastrow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'delete dummy column
Columns("B:B").Delete
Columns("O:O").Delete

End Sub

Sub checking_for_ERRORS()

Dim wsSrc As Worksheet
Dim selection1 As Range
Dim selection2 As Range
Dim lastrow As Long
Dim cell As Range
Dim colm As Range

Set wsSrc = Workbooks("BALPrint_Invoicing_ReferenceEntry.xlsx").Worksheets("DOC-E-006")

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

selection2.Copy
wsSrc.Range("AR2").PasteSpecial xlPasteValues

selection2.ClearContents

End Sub


