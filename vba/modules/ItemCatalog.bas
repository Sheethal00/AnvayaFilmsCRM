Attribute VB_Name = "ItemCatalog"
Option Explicit

' Default per-line-item prices for 03a_Quotation_LineItems (SPEC.md §3).
' Catalog lives in 19_Settings columns N:O (schema/19_settings.yaml
' item_catalog), pre-populated at build time -- not staff-entered like
' TeamList.

Private Const SETTINGS_SHEET As String = "19_Settings"
Private Const NAME_COLUMN As String = "N"
Private Const PRICE_COLUMN As String = "O"

' Returns the catalog's default price for itemName, or Empty if the item
' isn't in the catalog -- callers should treat Empty as "no suggestion,"
' not "price is zero."
Public Function LookupDefaultPrice(itemName As String) As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)

    Dim found As Range
    Set found = ws.Columns(NAME_COLUMN & ":" & NAME_COLUMN).Find( _
        What:=itemName, LookAt:=xlWhole)

    If found Is Nothing Then
        LookupDefaultPrice = Empty
    Else
        LookupDefaultPrice = ws.Cells(found.Row, PRICE_COLUMN).Value
    End If
End Function
