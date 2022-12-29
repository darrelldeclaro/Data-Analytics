Attribute VB_Name = "emea_MFC"
Public Sub EMEA_ManualFileCreation()

Dim wsSrc1 As Worksheet
Dim wsSrc2 As Worksheet
Dim wsDst As Worksheet
Dim srclastrow1 As Long
Dim srclastrow2 As Long
Dim dstlastrow As Long

Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\EMEA - Manual File Creation\Extracted Raw Data\EMEA MFC - 1.xlsx"
Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\EMEA - Manual File Creation\Extracted Raw Data\EMEA MFC - 2.xlsx"
Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\EMEA - Manual File Creation\MANUAL_FILE_COMBINED RAW (PBI).xlsx"

Set wsSrc1 = Workbooks("EMEA MFC - 1.xlsx").Worksheets("Sheet1")
Set wsSrc2 = Workbooks("EMEA MFC - 2.xlsx").Worksheets("Sheet1")
Set wsDst = Workbooks("MANUAL_FILE_COMBINED RAW (PBI)").Worksheets("MANUAL_FILE_COMBINED")

'copy and paste value from emea_mfc-1
srclastrow1 = wsSrc1.Cells(wsSrc1.Rows.Count, "A").End(xlUp).row

dstlastrow = wsDst.Cells(wsDst.Rows.Count, "A").End(xlUp).Offset(1).row

wsSrc1.Range("A2:T" & srclastrow1).Copy
wsDst.Range("A" & dstlastrow).PasteSpecial xlPasteValues
Application.CutCopyMode = False

'copy and paste values from emea_mfc-2
srclastrow2 = wsSrc2.Cells(wsSrc2.Rows.Count, "A").End(xlUp).row

dstlastrow = wsDst.Cells(wsDst.Rows.Count, "A").End(xlUp).Offset(1).row

wsSrc2.Range("A2:T" & srclastrow2).Copy
wsDst.Range("A" & dstlastrow).PasteSpecial xlPasteValues
Application.CutCopyMode = False

'closing extracted data
Workbooks("MANUAL_FILE_COMBINED RAW (PBI)").Activate
Workbooks("EMEA MFC - 1.xlsx").Close
Workbooks("EMEA MFC - 2.xlsx").Close


removingDuplicatesOpt_2
remove_selected_column


Application.ScreenUpdating = False
Columns("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
Application.ScreenUpdating = True

MsgBox ("Check columns: P - no unnecessary values Q - numerical values only")

End Sub


Sub removingDuplicatesOpt_2()

Dim selection1 As Range
Dim selection2 As Range
Dim selection3 As Range
Dim selection4 As Range
Dim lastrow As Long
Dim lastrowC As Long

lastrow = Cells(Rows.Count, 1).End(xlUp).row

'created 2 dummy column
Range("B:C").Insert
Range("B1") = "ord"
Range("C1") = "xty"
Range("B2") = 1
Range("B3") = 2


'column ORD
Set selection1 = ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Range("B2:B3")
Set selection2 = ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Range("B2:B" & lastrow)

selection1.autofill Destination:=selection2

Range("A1:V1").Select
Selection.AutoFilter

'Sort ticket number (smallest to largest) "Ascending"
ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Sort.SortFields.add Key:=Range("A1"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

With ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Sort
    .SetRange Range("A1:V" & lastrow)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With




'column XTY
Range("C2").FormulaR1C1 = "=IF(RC[-2]=R[1]C[-2],1,""x"")"

Set selection3 = ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Range("C2")
Set selection4 = ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Range("C2:C" & lastrow)

selection3.autofill Destination:=selection4

selection4.Copy
selection4.PasteSpecial xlPasteValues

'Ascending
ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Sort.SortFields.add Key:=Range("C1"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

With ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Sort
    .SetRange Range("A1:V" & lastrow)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With


'filtering
ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Range("C2").AutoFilter field:=3, Criteria1:=1



'Delete duplicate Rows
lastrowC = Cells(Rows.Count, 3).End(xlUp).row

Range(Cells(2, "C"), Cells(lastrowC, "C")).SpecialCells(xlCellTypeVisible).EntireRow.Delete
  
'Clear Filter
ActiveSheet.ShowAllData
    
    
'Ascending
ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Sort.SortFields.add Key:=Range("B1"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

With ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Sort
    .SetRange Range("A1:V" & lastrow)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With


'delete dummy columns
Range("B:C").Delete

End Sub





Sub remove_selected_column()

Dim lastrow As Long

lastrow = Cells(Rows.Count, 1).End(xlUp).row

Range("F2:F" & lastrow).ClearContents
Range("G2:G" & lastrow).ClearContents
Range("H2:H" & lastrow).ClearContents
Range("J2:J" & lastrow).ClearContents
Range("K2:K" & lastrow).ClearContents
Range("M2:M" & lastrow).ClearContents

End Sub

