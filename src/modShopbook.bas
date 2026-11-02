Attribute VB_Name = "modShopbook"
Option Explicit

' =============================================================================
'  Small Business Shopbook - day-to-day macros
'
'  Every routine works through the Excel table (ListObject) API and finds
'  columns by header name, so rows always land inside the tables and inserting
'  or moving a column can't break a macro.
'
'  Sheets / tables:
'    Sales            tblSales         one row per sale line
'    Supplier Orders  tblOrders        one row per order line, grouped by Order No.
'    Cashflow         tblTill          till summary (formulas)
'                     tblWithdrawals   cash taken out of the till
'    Catalogue        tblCatalogue     Brand > Size > Product, cost, price, stock
' =============================================================================

Private Const SH_SALES As String = "Sales"
Private Const SH_ORDERS As String = "Supplier Orders"
Private Const SH_CASH As String = "Cashflow"
Private Const SH_DASH As String = "Dashboard"
Private Const SH_CAT As String = "Catalogue"

Public Const WRITE_OFF As String = "WRITE-OFF"
Public Const OWNER_DRAWINGS As String = "OWNER DRAWINGS"

' ----------------------------------------------------------------- helpers

Private Function Tbl(ByVal sheetName As String, ByVal tableName As String) As ListObject
    Set Tbl = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
End Function

Private Function ColIx(ByVal t As ListObject, ByVal header As String) As Long
    ColIx = t.ListColumns(header).Index
End Function

Private Function Cell(ByVal r As ListRow, ByVal header As String) As Range
    Set Cell = r.Range.Cells(1, ColIx(r.Parent, header))
End Function

' Next number in a numeric ID column (1 for an empty table).
Private Function NextNumber(ByVal t As ListObject, ByVal header As String) As Long
    If t.ListRows.Count = 0 Then
        NextNumber = 1
    Else
        NextNumber = Application.WorksheetFunction.Max(t.ListColumns(header).DataBodyRange) + 1
    End If
End Function

Private Function LastRow(ByVal t As ListObject) As ListRow
    If t.ListRows.Count > 0 Then Set LastRow = t.ListRows(t.ListRows.Count)
End Function

Private Sub SelectCell(ByVal Target As Range)
    Application.GoTo Target
End Sub

' Delete the last data row of a table after confirmation. Header-safe: it only
' ever touches ListRows, never worksheet rows.
Private Sub DeleteLast(ByVal t As ListObject, ByVal idHeader As String, ByVal what As String)
    Dim r As ListRow
    Set r = LastRow(t)
    If r Is Nothing Then
        MsgBox "There are no " & what & "s to delete.", vbExclamation, "Nothing to delete"
        Exit Sub
    End If
    Dim label As String
    If Len(idHeader) > 0 Then label = " " & Cell(r, idHeader).Text
    If MsgBox("Delete the last " & what & label & "?", vbYesNo + vbQuestion, "Delete last " & what) = vbYes Then
        r.Delete
    End If
End Sub

' ----------------------------------------------------------------- sales

Public Sub AddSale()
    Dim t As ListObject, r As ListRow, inv As Long
    Set t = Tbl(SH_SALES, "tblSales")
    inv = NextNumber(t, "Invoice")
    Set r = t.ListRows.Add(AlwaysInsert:=True)
    Cell(r, "Invoice").Value = inv
    Cell(r, "Date").Value = Date
    SelectCell Cell(r, "Customer")
End Sub

Public Sub DeleteLastSale()
    DeleteLast Tbl(SH_SALES, "tblSales"), "Invoice", "sale"
End Sub

' ----------------------------------------------------------------- supplier orders

' Start a new supplier order: new Order No., today's date, PENDING.
Public Sub AddOrder()
    Dim t As ListObject, r As ListRow, n As Long
    Set t = Tbl(SH_ORDERS, "tblOrders")
    n = NextNumber(t, "Order No.")
    Set r = t.ListRows.Add(AlwaysInsert:=True)
    Cell(r, "Order No.").Value = n
    Cell(r, "Date").Value = Date
    Cell(r, "Status").Value = "PENDING"
    SelectCell Cell(r, "Brand")
End Sub

' Add another line to the most recent order (same number, date and status).
Public Sub AddOrderLine()
    Dim t As ListObject, prev As ListRow, r As ListRow
    Set t = Tbl(SH_ORDERS, "tblOrders")
    Set prev = LastRow(t)
    If prev Is Nothing Then
        AddOrder
        Exit Sub
    End If
    Set r = t.ListRows.Add(AlwaysInsert:=True)
    Cell(r, "Order No.").Value = Cell(prev, "Order No.").Value
    Cell(r, "Date").Value = Cell(prev, "Date").Value
    Cell(r, "Status").Value = Cell(prev, "Status").Value
    SelectCell Cell(r, "Brand")
End Sub

Public Sub DeleteLastOrderLine()
    DeleteLast Tbl(SH_ORDERS, "tblOrders"), "Order No.", "order line from order"
End Sub

' ----------------------------------------------------------------- cashflow

' Till Before / Till After are table formulas: cash sales up to the withdrawal
' date minus every earlier withdrawal. Nothing is stamped as a value, so
' editing an old row re-flows every balance after it.
Public Sub AddWithdrawal()
    Dim t As ListObject, r As ListRow
    Set t = Tbl(SH_CASH, "tblWithdrawals")
    Set r = t.ListRows.Add(AlwaysInsert:=True)
    Cell(r, "Date").Value = Date
    SelectCell Cell(r, "Cash Requested")
End Sub

Public Sub DeleteLastWithdrawal()
    DeleteLast Tbl(SH_CASH, "tblWithdrawals"), "Date", "withdrawal dated"
End Sub

' ----------------------------------------------------------------- dashboard

Public Sub RefreshDashboard()
    Dim pt As PivotTable
    Application.CalculateFull
    For Each pt In ThisWorkbook.Worksheets(SH_DASH).PivotTables
        pt.PivotCache.Refresh
    Next pt
End Sub

' ----------------------------------------------------------------- edit hooks
' Called from ThisWorkbook.Workbook_SheetChange (events are already disabled).

Public Sub OnTableEdit(ByVal t As ListObject, ByVal changed As Range)
    Dim isSales As Boolean
    isSales = (t.Name = "tblSales")

    Dim cBrand As Long, cSize As Long, cProduct As Long, cQty As Long, cCustomer As Long
    cBrand = ColIx(t, "Brand"): cSize = ColIx(t, "Size"): cProduct = ColIx(t, "Product")
    cQty = ColIx(t, "Quantity")
    If isSales Then cCustomer = ColIx(t, "Customer")

    Dim singleCell As Boolean
    singleCell = (changed.Cells.Count = 1)

    ' Brand or Size changed by hand -> the cells to its right are now invalid.
    If singleCell Then
        Dim c As Long
        c = changed.Column - t.Range.Column + 1
        If c = cBrand Then
            RowCell(t, changed, cSize).ClearContents
            RowCell(t, changed, cProduct).ClearContents
        ElseIf c = cSize Then
            RowCell(t, changed, cProduct).ClearContents
        End If
    End If

    Dim area As Range, rw As Range
    For Each area In changed.Areas
        For Each rw In area.Rows
            Dim touched As Range
            Set touched = Intersect(rw.EntireRow, t.DataBodyRange)
            If Not touched Is Nothing Then
                If Touches(t, rw, cBrand, cSize, cProduct) Then StampPrices t, touched, isSales
                If isSales Then
                    If Touches(t, rw, cCustomer) Then ForceOtherMethod t, touched
                    If singleCell And Touches(t, rw, cProduct, cQty) Then WarnIfOversold t, touched
                End If
            End If
        Next rw
    Next area
End Sub

Private Function RowCell(ByVal t As ListObject, ByVal anyCellInRow As Range, ByVal colIndex As Long) As Range
    Set RowCell = t.Range.Worksheet.Cells(anyCellInRow.Row, t.Range.Column + colIndex - 1)
End Function

Private Function Touches(ByVal t As ListObject, ByVal rw As Range, ParamArray cols() As Variant) As Boolean
    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        If Not Intersect(rw, t.ListColumns(cols(i)).DataBodyRange) Is Nothing Then
            Touches = True
            Exit Function
        End If
    Next i
End Function

Private Function RowVal(ByVal t As ListObject, ByVal rowRange As Range, ByVal header As String) As Variant
    RowVal = rowRange.Cells(1, ColIx(t, header)).Value
End Function

' Copies Unit Price / Unit Cost from the catalogue at the moment the product is
' picked. Past sales keep the price they were sold at even if the catalogue
' price changes later.
Private Sub StampPrices(ByVal t As ListObject, ByVal rowRange As Range, ByVal isSales As Boolean)
    Dim cat As ListObject, i As Long
    Set cat = Tbl(SH_CAT, "tblCatalogue")
    i = CatalogueIndex(cat, RowVal(t, rowRange, "Brand"), RowVal(t, rowRange, "Size"), RowVal(t, rowRange, "Product"))

    Dim costCell As Range
    Set costCell = rowRange.Cells(1, ColIx(t, "Unit Cost"))
    If i = 0 Then
        costCell.ClearContents
        If isSales Then rowRange.Cells(1, ColIx(t, "Unit Price")).ClearContents
    Else
        costCell.Value = cat.ListColumns("Cost").DataBodyRange.Cells(i, 1).Value
        If isSales Then
            rowRange.Cells(1, ColIx(t, "Unit Price")).Value = cat.ListColumns("Price").DataBodyRange.Cells(i, 1).Value
        End If
    End If
End Sub

' Write-offs and owner drawings never bring in money.
Private Sub ForceOtherMethod(ByVal t As ListObject, ByVal rowRange As Range)
    Dim who As String
    who = UCase$(Trim$(CStr(RowVal(t, rowRange, "Customer"))))
    If who = WRITE_OFF Or who = OWNER_DRAWINGS Then
        rowRange.Cells(1, ColIx(t, "Method")).Value = "OTHER"
    End If
End Sub

' Warns (does not block) when a sale line takes a product below zero stock.
Private Sub WarnIfOversold(ByVal t As ListObject, ByVal rowRange As Range)
    Dim b As Variant, s As Variant, p As Variant, remaining As Double
    b = RowVal(t, rowRange, "Brand"): s = RowVal(t, rowRange, "Size"): p = RowVal(t, rowRange, "Product")
    If Len(b) = 0 Or Len(s) = 0 Or Len(p) = 0 Then Exit Sub
    remaining = StockOf(b, s, p)
    If remaining < 0 Then
        MsgBox "Only " & (remaining + Val(RowVal(t, rowRange, "Quantity"))) & " of " & b & " " & s & " " & p & _
               " were in stock before this sale." & vbCrLf & _
               "Check the quantity, or mark the supplier order RECEIVED first.", _
               vbExclamation, "Not enough stock"
    End If
End Sub

' Live stock for one product: RECEIVED order quantity minus everything sold,
' written off or taken as owner drawings. Same rule as tblCatalogue[In Stock],
' computed here so it doesn't depend on recalculation order.
Public Function StockOf(ByVal brand As Variant, ByVal size As Variant, ByVal product As Variant) As Double
    Dim o As ListObject, sl As ListObject, wf As WorksheetFunction
    Set o = Tbl(SH_ORDERS, "tblOrders")
    Set sl = Tbl(SH_SALES, "tblSales")
    Set wf = Application.WorksheetFunction
    Dim received As Double, sold As Double
    If o.ListRows.Count > 0 Then
        received = wf.SumIfs(o.ListColumns("Quantity").DataBodyRange, _
            o.ListColumns("Brand").DataBodyRange, brand, o.ListColumns("Size").DataBodyRange, size, _
            o.ListColumns("Product").DataBodyRange, product, o.ListColumns("Status").DataBodyRange, "RECEIVED")
    End If
    If sl.ListRows.Count > 0 Then
        sold = wf.SumIfs(sl.ListColumns("Quantity").DataBodyRange, _
            sl.ListColumns("Brand").DataBodyRange, brand, sl.ListColumns("Size").DataBodyRange, size, _
            sl.ListColumns("Product").DataBodyRange, product)
    End If
    StockOf = received - sold
End Function

' 1-based row in tblCatalogue for Brand|Size|Product, or 0 if not found.
Private Function CatalogueIndex(ByVal cat As ListObject, ByVal brand As Variant, ByVal size As Variant, ByVal product As Variant) As Long
    If cat.ListRows.Count = 0 Then Exit Function
    Dim vb As Variant, vs As Variant, vp As Variant, i As Long
    vb = ColumnValues(cat, "Brand")
    vs = ColumnValues(cat, "Size")
    vp = ColumnValues(cat, "Product")
    For i = 1 To UBound(vb, 1)
        If CStr(vb(i, 1)) = CStr(brand) And CStr(vs(i, 1)) = CStr(size) And CStr(vp(i, 1)) = CStr(product) Then
            CatalogueIndex = i
            Exit Function
        End If
    Next i
End Function

' A table column as a 2-D (n x 1) array, even when the table has one row.
Private Function ColumnValues(ByVal t As ListObject, ByVal header As String) As Variant
    Dim rng As Range, v() As Variant
    Set rng = t.ListColumns(header).DataBodyRange
    If rng.Cells.Count = 1 Then
        ReDim v(1 To 1, 1 To 1)
        v(1, 1) = rng.Value
        ColumnValues = v
    Else
        ColumnValues = rng.Value
    End If
End Function
