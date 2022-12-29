Attribute VB_Name = "SCA_Other_tasks"
Public Sub SCA_Other_tasks()

Dim src1 As Worksheet
Dim src2 As Worksheet
Dim dst1 As Worksheet
Dim dst2 As Worksheet
Dim srclastrow1 As Long
Dim srclastrow2 As Long
Dim dstlastrow1 As Long
Dim dstlastrow2 As Long
Dim selection1 As Range
Dim selection2 As Range
Dim location As Range

Workbooks.Open ("C:\Users\darrell.declaro\Documents\Automation ver1.0\SCA - Other Tasks\Extracted from BO\SCA - Other Task.xlsx")
Workbooks.Open ("C:\Users\darrell.declaro\Documents\Automation ver1.0\SCA - Other Tasks\SCA - Other Task (Template).xlsx")
Workbooks.Open ("C:\Users\darrell.declaro\Documents\Automation ver1.0\SCA - Other Tasks\SCA - Other Task (BPI).xlsx")

Set src1 = Workbooks("SCA - Other Task.xlsx").Worksheets(1)
Set src2 = Workbooks("SCA - Other Task (Template).xlsx").Worksheets(1)
Set dst1 = Workbooks("SCA - Other Task (Template).xlsx").Worksheets(1)
Set dst2 = Workbooks("SCA - Other Task (BPI).xlsx").Worksheets(1)

srclastrow1 = src1.Cells(src1.Rows.Count, 1).End(xlUp).row
srclastrow2 = src2.Cells(src2.Rows.Count, 1).End(xlUp).row
dstlastrow1 = dst1.Cells(dst1.Rows.Count, 1).End(xlUp).row
dstlastrow2 = dst2.Cells(dst2.Rows.Count, 1).End(xlUp).row


'convert to number
With src1.Range("M2:M" & srclastrow1)
    .NumberFormat = "General"
    .Value = .Value
End With

'to template
dst1.Range("A2:V" & dstlastrow1).ClearContents
dst1.Range("W5:X" & dstlastrow1).ClearContents

src1.Range("A3:V" & srclastrow1).Copy
dst1.Range("A2").PasteSpecial xlPasteValues
Application.CutCopyMode = False

dst1.Range("W2:X2").Copy
dst1.Range("W2:X" & dstlastrow1).PasteSpecial xlPasteFormulas

Workbooks("SCA - Other Task.xlsx").Close savechanges:=False


'to pbi
dst2.Range("A1").CurrentRegion.ClearContents

src2.Range("A3").CurrentRegion.Copy

dst2.Range("A1").PasteSpecial xlPasteValues
Application.CutCopyMode = False


'convert to number
With dst2.Range("A2:A" & dstlastrow2)
    .NumberFormat = "dd/mm/yyyy hh:mm:ss"
    .Value = .Value
End With

With dst2.Range("I2:I" & dstlastrow2)
    .NumberFormat = "dd/mm/yyyy hh:mm:ss"
    .Value = .Value
End With

With dst2.Range("S2:S" & dstlastrow2)
    .NumberFormat = "dd/mm/yyyy hh:mm:ss"
    .Value = .Value
End With



Workbooks("SCA - Other Task (Template).xlsx").Close savechanges:=True
Workbooks("SCA - Other Task (BPI).xlsx").Close savechanges:=True

End Sub
