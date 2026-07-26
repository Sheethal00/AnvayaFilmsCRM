Option Explicit

' Merged into the 03a_Quotation_LineItems worksheet's built-in code
' module. See vba/classes/02_CRM_Leads.cls header note for how this gets
' wired in.
'
' Column layout: A QuoteID, B ItemDescription, C Qty, D UnitPrice,
' E LineTotal.

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub
    If Target.Column <> 2 Then Exit Sub ' only react to ItemDescription
    If Target.Value = "" Then Exit Sub

    Dim priceCell As Range
    Set priceCell = Me.Cells(Target.Row, 4) ' UnitPrice
    If priceCell.Value <> "" Then Exit Sub ' never overwrite a custom price

    Dim defaultPrice As Variant
    defaultPrice = ItemCatalog.LookupDefaultPrice(Target.Value)
    If IsEmpty(defaultPrice) Then Exit Sub ' item not in the catalog

    On Error GoTo CleanUp
    Application.EnableEvents = False
    priceCell.Value = defaultPrice
CleanUp:
    Application.EnableEvents = True
End Sub
