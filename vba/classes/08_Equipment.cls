Option Explicit

' Merged into the 08_Equipment worksheet's built-in code module.
' See vba/classes/02_CRM_Leads.cls header note for how this gets wired in.
'
' Column layout: A EquipmentID, B Category, C Model, D SerialNumber,
' E PurchaseDate, F Status, G CurrentBookingID, H MaintenanceNotes.

Private Const COL_EQUIPMENT_ID As Long = 1
Private Const COL_STATUS As Long = 6
Private Const COL_CURRENT_BOOKING_ID As Long = 7

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub

    On Error GoTo CleanUp

    IDGenerator.AutoFillID Me, Target.Row, COL_EQUIPMENT_ID, "LastEquipmentID"

    If Target.Column = COL_STATUS Or Target.Column = COL_CURRENT_BOOKING_ID Then
        CheckAndRevertIfConflict Target.Row
    End If

CleanUp:
    Application.EnableEvents = True
End Sub

Private Sub CheckAndRevertIfConflict(r As Long)
    Dim status As String, bookingID As String, equipmentID As String

    status = Me.Cells(r, COL_STATUS).Value
    bookingID = Me.Cells(r, COL_CURRENT_BOOKING_ID).Value
    equipmentID = Me.Cells(r, COL_EQUIPMENT_ID).Value

    If status <> "Assigned" Or bookingID = "" Or equipmentID = "" Then Exit Sub

    If Not AvailabilityCheck.CheckEquipmentAvailability(equipmentID, bookingID) Then
        Application.EnableEvents = False
        Me.Cells(r, COL_CURRENT_BOOKING_ID).ClearContents
        Application.EnableEvents = True
    End If
End Sub
