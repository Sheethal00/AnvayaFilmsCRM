Attribute VB_Name = "LeadConversion"
Option Explicit

' SPEC.md §3 "02_CRM_Leads" automation: on Status -> Won, create the
' matching row in 04_Client_Master and write ConvertedClientID back onto
' the lead. Prevents duplicate manual entry between Lead Management and
' Customer Database (the failure mode the original spec's parallel,
' unconnected sheets invited).

Private Const LEADS_SHEET As String = "02_CRM_Leads"
Private Const CLIENTS_SHEET As String = "04_Client_Master"

' Column layout (schema/02_crm_leads.yaml): A LeadID, B DateReceived,
' C ClientName, D Phone, E Email, F Source, G EventType, H EventDate,
' I Budget, J Status, K LostReason, L AssignedTo, M NextFollowUpDate,
' N Notes, O ConvertedClientID.
Private Const COL_CLIENT_NAME As Long = 3
Private Const COL_PHONE As Long = 4
Private Const COL_EMAIL As Long = 5
Private Const COL_SOURCE As Long = 6
Private Const COL_CONVERTED_CLIENT_ID As Long = 15

' Column layout (schema/04_client_master.yaml): A ClientID,
' B PrimaryContactName, C Phone, D Email, E Address, F FamilyContacts,
' G Source, H ClientSince, I TotalLifetimeValue, J Notes.

Public Sub ConvertLeadToClient(leadRow As Long)
    Dim wsLead As Worksheet, wsClient As Worksheet
    Set wsLead = ThisWorkbook.Worksheets(LEADS_SHEET)
    Set wsClient = ThisWorkbook.Worksheets(CLIENTS_SHEET)

    If wsLead.Cells(leadRow, COL_CONVERTED_CLIENT_ID).Value <> "" Then
        Exit Sub ' already converted
    End If

    Dim newClientID As String
    newClientID = IDGenerator.GenerateID("LastClientID")

    Dim newRow As Long
    newRow = wsClient.Cells(wsClient.Rows.Count, "A").End(xlUp).Row + 1

    On Error GoTo CleanUp
    Application.EnableEvents = False

    wsClient.Cells(newRow, 1).Value = newClientID
    wsClient.Cells(newRow, 2).Value = wsLead.Cells(leadRow, COL_CLIENT_NAME).Value
    wsClient.Cells(newRow, 3).Value = wsLead.Cells(leadRow, COL_PHONE).Value
    wsClient.Cells(newRow, 4).Value = wsLead.Cells(leadRow, COL_EMAIL).Value
    wsClient.Cells(newRow, 7).Value = wsLead.Cells(leadRow, COL_SOURCE).Value

    wsLead.Cells(leadRow, COL_CONVERTED_CLIENT_ID).Value = newClientID

CleanUp:
    Application.EnableEvents = True
End Sub
