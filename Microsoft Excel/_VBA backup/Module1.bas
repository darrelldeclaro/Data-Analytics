Attribute VB_Name = "Module1"
Sub locate_and_paste()
'locating the cell value from source to destination, then paste
Set wsSrc = Workbooks("BALPrint_Invoicing_ReferenceEntry.xlsx").Worksheets("DOC-E-006")
Set wsDst = Workbooks("BAL_INV Volume (PBI).xlsx").Worksheets("BAL_INV")
Set tckId = wsSrc.Range("B2")
Set location = wsDst.Range("B:B").Find(tckId)

srclastrow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).row

wsSrc.Range("A2:AR" & srclastrow).Copy
wsDst.Range(location.Address).Offset(, -1).PasteSpecial xlPasteValuesAndNumberFormats
Application.CutCopyMode = False

End Sub
