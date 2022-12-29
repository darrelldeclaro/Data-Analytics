Attribute VB_Name = "TW_Delivery_Leg"
Public Sub TW_Delivery_Leg()

Dim src As Worksheet
Dim dst1 As Worksheet
Dim dst2 As Worksheet
Dim srclastrow As Long
Dim dstlastrow As Long
Dim lastrow As Long
Dim lastrow1 As Long


Workbooks.Open ("C:\Users\darrell.declaro\Documents\Automation ver1.0\ASP - Taiwan Delivery Leg\TW Delivery Leg - TEMPLATE.xlsx")
Workbooks.Open ("C:\Users\darrell.declaro\Documents\Automation ver1.0\ASP - Taiwan Delivery Leg\Extracted Raw Data\TW Delivery Leg (SSC_HER_Productivity) 2022.xlsx")


Set src = Workbooks("TW Delivery Leg (SSC_HER_Productivity) 2022.xlsx").Worksheets(1)
Set dst1 = Workbooks("TW Delivery Leg - TEMPLATE.xlsx").Worksheets(1)



srclastrow = src.Cells(src.Rows.Count, 1).End(xlUp).row


dst1.Range("A2:J" & 100000).ClearContents

src.Range("B5").CurrentRegion.Offset(1).Copy
dst1.Range("A2").PasteSpecial xlPasteValuesAndNumberFormats

lastrow = dst1.Cells(dst1.Rows.Count, 1).End(xlUp).row
dst1.Range("K2:AW2").Copy
dst1.Range("K2:AW" & lastrow).PasteSpecial xlPasteFormulas

Workbooks("TW Delivery Leg (SSC_HER_Productivity) 2022.xlsx").Close savechanges:=False

'custom data type format
With dst1.Range("A2:A" & lastrow)
    .NumberFormat = "general"
    .Value = .Value
End With

'filter cells that contains "test" or "tst" and delete entire row
dst1.Range("A2:AW" & lastrow).AutoFilter field:=2, Criteria1:=Array("*test*", "*tst*"), Operator:=xlFilterValues
lastrow1 = dst1.Cells(dst1.Rows.Count, 2).End(xlUp).row
dst1.Range(dst1.Cells(2, "B"), dst1.Cells(lastrow1, "B")).SpecialCells(xlCellTypeVisible).EntireRow.Delete
dst1.ShowAllData
Workbooks("TW Delivery Leg - TEMPLATE.xlsx").Save

cleanToPbi

End Sub


Sub cleanToPbi()

Dim src As Worksheet
Dim dst As Worksheet
Dim srclastrow As Long
Dim dstlastrow As Long

Workbooks.Open ("C:\Users\darrell.declaro\Documents\Automation ver1.0\ASP - Taiwan Delivery Leg\BUNK\TW Delivery Leg.xlsx")

Set src = Workbooks("TW Delivery Leg - TEMPLATE.xlsx").Worksheets(1)
Set dst = Workbooks("TW Delivery Leg.xlsx").Worksheets(1)

srclastrow = src.Cells(src.Rows.Count, 1).End(xlUp).row
dstlastrow = dst.Cells(dst.Rows.Count, 1).End(xlUp).row

dst.Range("A2:AW" & dstlastrow).Delete

src.Range("A2:AW" & srclastrow).Copy
dst.Range("A2").PasteSpecial xlPasteValuesAndNumberFormats
Application.CutCopyMode = False


'custom data type format
With dst.Range("G2:J" & dstlastrow)
    .NumberFormat = "dd/mm/yyyy hh:mm:ss"
    .Value = .Value
End With

With dst.Range("Q2:S" & dstlastrow)
    .NumberFormat = "dd/mm/yyyy hh:mm:ss"
    .Value = .Value
End With


Workbooks("TW Delivery Leg - TEMPLATE.xlsx").Close savechanges:=True
Workbooks("TW Delivery Leg.xlsx").Close savechanges:=True


End Sub
