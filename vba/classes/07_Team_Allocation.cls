Option Explicit

' Merged into the 07_Team_Allocation worksheet's built-in code module.
' See vba/classes/02_CRM_Leads.cls header note for how this gets wired in.
'
' Column layout: A AllocationID, B BookingID, C TeamMemberID, D Role,
' E EventDate, F Status.

Private Const COL_ALLOCATION_ID As Long = 1
Private Const COL_BOOKING_ID As Long = 2
Private Const COL_TEAM_MEMBER_ID As Long = 3
Private Const COL_EVENT_DATE As Long = 5

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub

    On Error GoTo CleanUp

    IDGenerator.AutoFillID Me, Target.Row, COL_ALLOCATION_ID, "LastAllocationID"

    ' Only re-check availability when the columns that define the
    ' conflict (who / when) are what just changed.
    If Target.Column = COL_TEAM_MEMBER_ID Or Target.Column = COL_EVENT_DATE Then
        CheckAndRevertIfConflict Target.Row
    End If

CleanUp:
    Application.EnableEvents = True
End Sub

Private Sub CheckAndRevertIfConflict(r As Long)
    Dim teamMemberID As String, bookingID As String
    Dim eventDate As Variant

    teamMemberID = Me.Cells(r, COL_TEAM_MEMBER_ID).Value
    bookingID = Me.Cells(r, COL_BOOKING_ID).Value
    eventDate = Me.Cells(r, COL_EVENT_DATE).Value

    If teamMemberID = "" Or bookingID = "" Or Not IsDate(eventDate) Then Exit Sub

    If Not AvailabilityCheck.CheckTeamAvailability(teamMemberID, CDate(eventDate), bookingID) Then
        ' Conflict: clear the cell that triggered this check rather than
        ' silently allowing a double-booking. MsgBox already shown inside
        ' CheckTeamAvailability.
        Application.EnableEvents = False
        Me.Cells(r, COL_TEAM_MEMBER_ID).ClearContents
        Application.EnableEvents = True
    End If
End Sub
