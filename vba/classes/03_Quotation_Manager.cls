Option Explicit

' Merged into the 03_Quotation_Manager worksheet's built-in code module.
' See vba/classes/02_CRM_Leads.cls header note for how this gets wired in.
'
' Column layout: A QuoteID, B LeadID, C DateIssued, D PackageSelected,
' E Subtotal, F GST%, G GSTAmount, H Total, I ValidUntil, J Status.

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub

    On Error GoTo CleanUp
    IDGenerator.AutoFillID Me, Target.Row, 1, "LastQuoteID"
CleanUp:
    Application.EnableEvents = True
End Sub
