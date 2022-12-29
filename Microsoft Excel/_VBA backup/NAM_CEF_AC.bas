Attribute VB_Name = "NAM_CEF_AC"
Public Sub NAM_CostcoExportFiles_AdditionalCheck()

Dim wsSrc As Worksheet
Dim wsDst As Worksheet
Dim srclastrow As Long
Dim dstlastrow As Long
Dim newLastRow As Long

Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\NAM - Costco Export Files - Additional Check\Extracted from RTQ\Results.xlsx"
Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\NAM - Costco Export Files - Additional Check\Costco Export Files - TEMPLATE (RT).xlsx"
Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\NAM - Costco Export Files - Additional Check\Costco Export Files (PBI data).xlsx"

'*****PATH RESULT TO TEMPLATE*****'
Set wsSrc = Workbooks("Results.xlsx").Worksheets("Sheet1")
Set wsDst = Workbooks("Costco Export Files - TEMPLATE (RT).xlsx").Worksheets("Template")

Workbooks("Costco Export Files - TEMPLATE (RT).xlsx").Activate

'copy until last row of the source'
srclastrow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).row

'determine the last row of the existing data'
dstlastrow = wsDst.Cells(wsDst.Rows.Count, "A").End(xlUp).Offset(1).row

'removing the existing data
Range("A2:P" & dstlastrow).ClearContents

'copy and paste'
wsSrc.Range("A2:P" & srclastrow).Copy _
    wsDst.Range("A2")

newLastRow = Cells(Rows.Count, 1).End(xlUp).row

'applying calculated fields from column Q-AA
Range("Q2:AA2").Copy
Range("Q2:AA" & newLastRow).PasteSpecial xlPasteFormulasAndNumberFormats
Application.CutCopyMode = False

'MsgBox ("Result to Template Complete")'
Workbooks("Results.xlsx").Close
Workbooks("Costco Export Files (PBI data).xlsx").Activate

'*****PATH TEMPLATE TO PBI******'
Set wsSrc = Workbooks("Costco Export Files - TEMPLATE (RT).xlsx").Worksheets("Template")
Set wsDst = Workbooks("Costco Export Files (PBI data).xlsx").Worksheets("NAM Costco EF - RT")

'copy until last row of the source'
ScrLastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).row

'find the first blank row of the destination'
dstlastrow = wsDst.Cells(wsDst.Rows.Count, "A").End(xlUp).Offset(1).row

'copy and paste'
wsSrc.Range("A2:AA" & srclastrow).Copy
wsDst.Range("A" & dstlastrow).PasteSpecial xlPasteValues
Application.CutCopyMode = False

Range("A2").Copy
Range("A2:A" & newLastRow).PasteSpecial xlPasteFormats
'MsgBox ("Template to PBI Complete")'
Workbooks("Costco Export Files - TEMPLATE (RT).xlsx").Close savechanges:=True


'MsgBox "Removing duplicates and highlighted fonts"

removingDuplicatesOpt1
remove_selected_column


'MsgBox "Column P - custom field {no. of transaction} - data should be numerical values ONLY"
checkingForErrors


Workbooks("Costco Export Files (PBI data).xlsx").Close savechanges:=True

'MsgBox ("PBI Completed and saved")'


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

Range("G1").End(xlDown).Select
Cells(, 23).Range(Selection, Selection.End(xlDown)) = "retain"

'adding index column
Range("B:B").Insert
Range("B2") = 1
Range("B3") = 2

Set selection1 = ActiveWorkbook.Worksheets("NAM Costco EF - RT").Range("B2:B3")
Set selection2 = ActiveWorkbook.Worksheets("NAM Costco EF - RT").Range("B2:B" & lastrow)

'autofill
selection1.autofill Destination:=selection2


'sort to descending
ActiveWorkbook.Worksheets("NAM Costco EF - RT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("NAM Costco EF - RT").Sort.SortFields.add Key:=Range("B1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("NAM Costco EF - RT").Sort
        .SetRange Range("A2:AD" & lastrow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
'remove duplicates
ActiveWorkbook.Worksheets("NAM Costco EF - RT").Range("A1:AD" & lastrow).removeDuplicates Columns:=1, Header:=xlYes


'sort to ascending
ActiveWorkbook.Worksheets("NAM Costco EF - RT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("NAM Costco EF - RT").Sort.SortFields.add Key:=Range("B1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("NAM Costco EF - RT").Sort
        .SetRange Range("A2:AD" & lastrow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'delete dummy column
Columns("B:B").Delete
Columns("ac:ac").Delete

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

Sub checkingForErrors()

Dim wsDst2 As Worksheet
Dim lastrow As Long
Dim colm As Range
Dim cell As Range

Set wsDst2 = Workbooks("Costco Export Files (PBI data).xlsx").Worksheets("NAM Costco EF - RT")

lastrow = wsDst2.Cells(wsDst2.Rows.Count, "P").End(xlUp).row

Set colm = wsDst2.Range("P2:P" & lastrow)
For Each cell In colm
    If cell > 1000 Then
        cell = 1
    ElseIf cell = "" Then
        cell = ""
    End If
Next cell



End Sub

