VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Every edit on Sales or Supplier Orders goes through modShopbook.OnTableEdit:
' it stamps Unit Price / Unit Cost from the catalogue, clears stale Size/Product
' picks, forces Method = OTHER for write-offs/drawings, and warns on overselling.
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim t As ListObject, hit As Range
    Select Case Sh.Name
        Case "Sales": Set t = Sh.ListObjects("tblSales")
        Case "Supplier Orders": Set t = Sh.ListObjects("tblOrders")
        Case Else: Exit Sub
    End Select
    If t.DataBodyRange Is Nothing Then Exit Sub
    Set hit = Intersect(Target, t.DataBodyRange)
    If hit Is Nothing Then Exit Sub

    On Error GoTo Done
    Application.EnableEvents = False
    modShopbook.OnTableEdit t, hit
Done:
    Application.EnableEvents = True
End Sub
