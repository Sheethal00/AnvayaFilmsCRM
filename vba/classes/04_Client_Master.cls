Option Explicit

' Merged into the 04_Client_Master worksheet's built-in code module.
' See vba/classes/02_CRM_Leads.cls header note for how this gets wired in.
'
' Column layout: A ClientID, B PrimaryContactName, C Phone, D Email,
' E Address, F FamilyContacts, G Source, H ClientSince,
' I TotalLifetimeValue, J Notes.
'
' Handles clients added directly on this sheet. Clients created via
' LeadConversion.ConvertLeadToClient already have ClientID set before
' this fires, so AutoFillID is a no-op for those rows.

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub

    On Error GoTo CleanUp
    IDGenerator.AutoFillID Me, Target.Row, 1, "LastClientID"
CleanUp:
    Application.EnableEvents = True
End Sub
