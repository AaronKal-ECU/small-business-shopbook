Attribute VB_Name = "modSetup"
Option Explicit

' =============================================================================
'  One-click setup: (re)creates the macro buttons and the buyer pivot.
'  Safe to run again - it deletes its own buttons/pivot before rebuilding.
' =============================================================================

Private Const BTN_H As Double = 26

' Office shape constants as plain values, so this module compiles without the
' "Microsoft Office Object Library" reference (it isn't always ticked on Mac).
Private Const SHAPE_ROUNDED_RECT As Long = 5    ' msoShapeRoundedRectangle
Private Const MSO_TRUE As Long = -1             ' msoTrue
Private Const MSO_FALSE As Long = 0             ' msoFalse
Private Const ANCHOR_MIDDLE As Long = 3         ' msoAnchorMiddle
Private Const ALIGN_CENTER As Long = 2          ' msoAlignCenter

Public Sub SetupShopbook()
    Application.ScreenUpdating = False
    CreateButtons
    BuildBuyerPivot
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets("Sales").Activate
    MsgBox "Buttons and the buyer pivot are ready.", vbInformation, "Shopbook setup"
End Sub

Public Sub CreateButtons()
    Dim green As Long, red As Long, blue As Long
    green = RGB(46, 125, 50): red = RGB(183, 28, 28): blue = RGB(31, 78, 121)

    With ThisWorkbook.Worksheets("Sales")
        AddButton .Range("A1"), "btnAddSale", "+ New Sale", "AddSale", 120, green
        AddButton .Range("C1"), "btnDeleteSale", "Delete Last Sale", "DeleteLastSale", 130, red
    End With
    With ThisWorkbook.Worksheets("Supplier Orders")
        AddButton .Range("A1"), "btnAddOrder", "+ New Order", "AddOrder", 120, green
        AddButton .Range("C1"), "btnAddOrderLine", "+ Add Line", "AddOrderLine", 110, blue
        AddButton .Range("E1"), "btnDeleteOrderLine", "Delete Last Line", "DeleteLastOrderLine", 130, red
    End With
    With ThisWorkbook.Worksheets("Cashflow")
        AddButton .Range("A4"), "btnAddWithdrawal", "+ Withdrawal", "AddWithdrawal", 120, green
        AddButton .Range("C4"), "btnDeleteWithdrawal", "Delete Last Withdrawal", "DeleteLastWithdrawal", 170, red
    End With
    With ThisWorkbook.Worksheets("Dashboard")
        AddButton .Range("H1"), "btnRefresh", "Refresh Dashboard", "RefreshDashboard", 150, blue
    End With
End Sub

Private Sub AddButton(ByVal anchor As Range, ByVal nm As String, ByVal caption As String, _
                      ByVal macroName As String, ByVal w As Double, ByVal colour As Long)
    Dim ws As Worksheet, s As Shape
    Set ws = anchor.Worksheet
    On Error Resume Next
    ws.Shapes(nm).Delete
    On Error GoTo 0

    Set s = ws.Shapes.AddShape(SHAPE_ROUNDED_RECT, anchor.Left + 4, anchor.Top + 4, w, BTN_H)
    s.Name = nm
    s.OnAction = macroName
    s.Placement = xlFreeFloating
    s.Fill.ForeColor.RGB = colour
    s.Line.Visible = MSO_FALSE
    With s.TextFrame2
        .VerticalAnchor = ANCHOR_MIDDLE
        .MarginLeft = 4: .MarginRight = 4: .MarginTop = 0: .MarginBottom = 0
        With .TextRange
            .Text = caption
            .Font.size = 11
            .Font.Bold = MSO_TRUE
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = ALIGN_CENTER
        End With
    End With
End Sub

' Top 10 customers by units (ties can show a few more rows), excluding Method = OTHER (write-offs and owner
' drawings), with what each customer spent.
Public Sub BuildBuyerPivot()
    Dim ws As Worksheet, pt As PivotTable, pc As PivotCache
    Set ws = ThisWorkbook.Worksheets("Dashboard")

    For Each pt In ws.PivotTables
        pt.TableRange2.Clear
    Next pt
    ws.Range("B23").ClearContents

    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="tblSales")
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("B25"), tableName:="ptBuyers")

    With pt
        .RowAxisLayout xlTabularRow
        .TableStyle2 = "PivotStyleLight16"
        .HasAutoFormat = False          ' refreshing must not resize the dashboard columns
        .ColumnGrand = False            ' a top-10 grand total would disagree with the KPI cards
        .RowGrand = False
        With .PivotFields("Method")
            .Orientation = xlPageField
            .EnableMultiplePageItems = True
            .PivotItems("OTHER").Visible = False
        End With
        .PivotFields("Customer").Orientation = xlRowField
        .AddDataField .PivotFields("Quantity"), "Units Bought", xlSum
        .AddDataField .PivotFields("Total"), "Spent", xlSum
        .DataFields("Spent").NumberFormat = "$#,##0.00"
        .PivotFields("Customer").AutoSort xlDescending, "Units Bought"
        .PivotFields("Customer").PivotFilters.Add2 Type:=xlTopCount, _
            DataField:=.DataFields("Units Bought"), Value1:=10
    End With
End Sub
