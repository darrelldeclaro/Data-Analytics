Attribute VB_Name = "EMEA_DEM_DET"
Public Sub EMEA_DEM_DET()

Dim wsSrc As Worksheet
Dim wsDst As Worksheet
Dim srclastrow As Long
Dim dstlastrow As Long

Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\EMEA - Demurrage and Detention\Extracted Raw Data\Results.xlsx"
Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\EMEA - Demurrage and Detention\DEMURRAGE_DETENTION Raw (PBI).xlsx"

'*****PATH RESULT TO TEMPLATE*****'
Set wsSrc = Workbooks("Results.xlsx").Worksheets("Sheet1")
Set wsDst = Workbooks("DEMURRAGE_DETENTION Raw (PBI).xlsx").Worksheets("DEMURRAGE_DETENTION")

'copy until last row of the source'
srclastrow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).row

'find the first blank row of the destination'
dstlastrow = wsDst.Cells(wsDst.Rows.Count, "A").End(xlUp).Offset(1).row

'copy and paste'
wsSrc.Range("A2:S" & srclastrow).Copy
    wsDst.Range("A" & dstlastrow).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
'MsgBox ("Result to PBI Complete")'
Workbooks("Results.xlsx").Close
Workbooks("DEMURRAGE_DETENTION Raw (PBI).xlsx").Activate

'MsgBox "Removing duplicates and highlighted fonts"

removingDuplicatesOpt1
remove_selected_column


Workbooks("DEMURRAGE_DETENTION Raw (PBI).xlsx").Save

MsgBox ("custom field {EMEA Country} - make sure that values are correct country code (2-letter)      *custom field {no. of transaction} - data should be numerical values ONLY")


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

Set selection1 = ActiveWorkbook.Worksheets(1).Range("B2:B3")
Set selection2 = ActiveWorkbook.Worksheets(1).Range("B2:B" & lastrow)

'autofill
selection1.autofill Destination:=selection2


'sort to descending
ActiveWorkbook.Worksheets(1).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(1).Sort.SortFields.add Key:=Range("B1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(1).Sort
        .SetRange Range("A2:AD" & lastrow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
'remove duplicates
ActiveWorkbook.Worksheets(1).Range("A1:AD" & lastrow).removeDuplicates Columns:=1, Header:=xlYes


'sort to ascending
ActiveWorkbook.Worksheets(1).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(1).Sort.SortFields.add Key:=Range("B1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(1).Sort
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






