Attribute VB_Name = "NAM_DO"
Public Sub NAM_Delivey_Order()

Dim wsSrc As Worksheet

'***** Initial Data Cleaning and Aggregation Process ****'
cleaningBO

'***** Second Phase Data Cleaning and Aggregation Process ****'
cleaningRTQ

'***** Third Phase - Clean data from RTQ to Complier ****'
complier1

'***** Third Phase - Clean data from complier to PBI ****'
complier_PBI

End Sub


Sub cleaningBO()

Dim src As Worksheet
Dim dst As Worksheet
Dim selection1 As Range
Dim selection2 As Range
Dim selection3 As Range
Dim selection4 As Range
Dim srclastrow As Long
Dim dstlastrow As Long

Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\NAM - Delivery Order\Raw Data\Delivery Order Report (Sep - Dec 2022).xlsx"
Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\NAM - Delivery Order\BO Template.xlsx"

Set src = Workbooks("Delivery Order Report (Sep - Dec 2022).xlsx").Worksheets(1)
Set dst = Workbooks("BO Template.xlsx").Worksheets(2)

srclastrow = src.Cells(src.Rows.Count, "A").End(xlUp).row
dstlastrow = dst.Cells(dst.Rows.Count, "B").End(xlUp).row

dst.Range("B2:O" & dstlastrow).ClearContents
dst.Range("A5:Y" & dstlastrow).ClearContents

src.Range("A3:N" & srclastrow).Copy
dst.Range("B2").PasteSpecial xlPasteValues
Application.CutCopyMode = False

'column A
lastrow = dst.Cells(Rows.Count, "B").End(xlUp).row
Set selection1 = dst.Range("A2")
Set selection2 = dst.Range("A2:A" & lastrow)

selection1.autofill Destination:=selection2

'column P-Y
Set selection3 = dst.Range("P2:Y2")
Set selection4 = dst.Range("P2:Y" & lastrow)

selection3.autofill Destination:=selection4

Workbooks("BO Template.xlsx").Save
Workbooks("Delivery Order Report (Sep - Dec 2022).xlsx").Close savechanges:=False

End Sub


Sub cleaningRTQ()

Dim src As Worksheet
Dim dst As Worksheet
Dim selection1 As Range
Dim selection2 As Range
Dim srclastrow As Long
Dim dstlastrow As Long

Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\NAM - Delivery Order\Raw Data\Results.xlsx"
Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\NAM - Delivery Order\RT Template.xlsx"


Set src = Workbooks("Results.xlsx").Worksheets(1)
Set dst = Workbooks("RT Template.xlsx").Worksheets(1)

srclastrow = src.Cells(src.Rows.Count, "A").End(xlUp).row
dstlastrow = dst.Cells(dst.Rows.Count, "A").End(xlUp).Offset(1).row

src.Range("A2:O" & srclastrow).Copy
dst.Range("A" & dstlastrow).PasteSpecial xlPasteValues
Application.CutCopyMode = False

'column P-AV
lastrow = dst.Cells(Rows.Count, "B").End(xlUp).row
Set selection1 = dst.Range("P2:AV2")
Set selection2 = dst.Range("P2:AV" & lastrow)

selection1.autofill Destination:=selection2

remove_selected_column

Workbooks("RT Template.xlsx").Save
Workbooks("Results.xlsx").Close savechanges:=False

End Sub


Sub complier1()

Dim src As Worksheet
Dim dst As Worksheet
Dim srclastrow As Long
Dim dstlastrow As Long

Set src = Workbooks("BO Template.xlsx").Worksheets(2)
Set dst = Workbooks("RT Template.xlsx").Worksheets(2)

Set tckId = src.Range("A2")
Set location = dst.Range("A:A").Find(tckId)

srclastrow = src.Cells(src.Rows.Count, "A").End(xlUp).row
dstlastrow = dst.Cells(dst.Rows.Count, "A").End(xlUp).row

src.Range("A2:Y" & srclastrow).Copy
dst.Range(location.Address).PasteSpecial xlPasteValues
Application.CutCopyMode = False

Workbooks("BO Template.xlsx").Close savechanges:=True

End Sub

Sub complier_PBI()

Dim src As Worksheet
Dim dst As Worksheet
Dim selection1 As Range
Dim selection2 As Range
Dim srclastrow As Long
Dim dstlastrow As Long

Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\NAM - Delivery Order\NAM Delivery Order (PBI).xlsx"

Set src = Workbooks("RT Template.xlsx").Worksheets(1)
Set dst = Workbooks("NAM Delivery Order (PBI).xlsx").Worksheets(1)

srclastrow = src.Cells(src.Rows.Count, "A").End(xlUp).row
dstlastrow = dst.Cells(dst.Rows.Count, "A").End(xlUp).Offset(1).row

src.Range("A2:AV" & srclastrow).Copy
dst.Range("A" & dstlastrow).PasteSpecial xlPasteValues
Application.CutCopyMode = False

removingDuplicatesOpt_2

Workbooks("RT Template.xlsx").Close savechanges:=True
Workbooks("NAM Delivery Order (PBI).xlsx").Close savechanges:=True

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
Set selection1 = ActiveWorkbook.Worksheets(1).Range("B2:B3")
Set selection2 = ActiveWorkbook.Worksheets(1).Range("B2:B" & lastrow)

selection1.autofill Destination:=selection2

Range("A1:AV1").Select


'Sort ticket number (smallest to largest) "Ascending"
ActiveWorkbook.Worksheets(1).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(1).Sort.SortFields.add Key:=Range("A1"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

With ActiveWorkbook.Worksheets(1).Sort
    .SetRange Range("A1:AV" & lastrow)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With




'column XTY
Range("C2").FormulaR1C1 = "=IF(RC[-2]=R[1]C[-2],1,""x"")"

Set selection3 = ActiveWorkbook.Worksheets(1).Range("C2")
Set selection4 = ActiveWorkbook.Worksheets(1).Range("C2:C" & lastrow)

selection3.autofill Destination:=selection4

selection4.Copy
selection4.PasteSpecial xlPasteValues

'Ascending
ActiveWorkbook.Worksheets(1).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(1).Sort.SortFields.add Key:=Range("C1"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

With ActiveWorkbook.Worksheets(1).Sort
    .SetRange Range("A1:AV" & lastrow)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With


'filtering
ActiveWorkbook.Worksheets(1).Range("C2").AutoFilter field:=3, Criteria1:=1



'Delete duplicate Rows
If Range("C2") = 1 Then
    lastrowC = Cells(Rows.Count, 3).End(xlUp).row
    
    Range(Cells(2, "C"), Cells(lastrowC, "C")).SpecialCells(xlCellTypeVisible).EntireRow.Delete
      
    'Clear Filter
    ActiveSheet.ShowAllData
ElseIf Range("C2") = "x" Then
    ActiveSheet.ShowAllData
End If
    
'Ascending
ActiveWorkbook.Worksheets(1).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(1).Sort.SortFields.add Key:=Range("B1"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

With ActiveWorkbook.Worksheets(1).Sort
    .SetRange Range("A1:AV" & lastrow)
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

Range("G2:G" & lastrow).ClearContents
Range("H2:H" & lastrow).ClearContents
Range("M2:M" & lastrow).ClearContents

End Sub






