Option Explicit

' Merged into the 09_Shoot_Checklist worksheet's built-in code module.
' See vba/classes/02_CRM_Leads.cls header note for how this gets wired in.
'
' Column layout: A BookingID, B ChecklistPhase, C ItemName, D Checked,
' E CheckedBy, F CheckedAt, G Approved, H ApprovedBy, I ApprovedAt.
'
' Two-stage workflow: the allocated crew member checks an item off, then
' someone else approves it. ApprovedBy is auto-recorded from
' Environ("Username") -- like UserRolePermissions (CLAUDE.md caveat),
' this is informational, not real access control: it records who was
' logged into Windows at the time, it doesn't restrict who can tick the
' box.

Private Const COL_BOOKING_ID As Long = 1
Private Const COL_CHECKED As Long = 4
Private Const COL_CHECKED_BY As Long = 5
Private Const COL_CHECKED_AT As Long = 6
Private Const COL_APPROVED As Long = 7
Private Const COL_APPROVED_BY As Long = 8
Private Const COL_APPROVED_AT As Long = 9

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub

    On Error GoTo CleanUp
    Application.EnableEvents = False

    Select Case Target.Column
        Case COL_CHECKED
            HandleCheckedChange Target.Row

        Case COL_CHECKED_BY
            EnforceCheckedByIsAllocated Target.Row

        Case COL_APPROVED
            HandleApprovedChange Target.Row
    End Select

CleanUp:
    Application.EnableEvents = True
End Sub

Private Sub HandleCheckedChange(r As Long)
    If Me.Cells(r, COL_CHECKED).Value = True Then
        Me.Cells(r, COL_CHECKED_AT).Value = Now
    Else
        Me.Cells(r, COL_CHECKED_AT).ClearContents
        ' Unchecking revokes any approval too -- "approved" implies "checked".
        Me.Cells(r, COL_APPROVED).Value = False
        Me.Cells(r, COL_APPROVED_BY).ClearContents
        Me.Cells(r, COL_APPROVED_AT).ClearContents
    End If
End Sub

Private Sub EnforceCheckedByIsAllocated(r As Long)
    Dim memberName As String, bookingID As String
    memberName = Me.Cells(r, COL_CHECKED_BY).Value
    bookingID = Me.Cells(r, COL_BOOKING_ID).Value

    If memberName = "" Or bookingID = "" Then Exit Sub

    If Not ChecklistWorkflow.IsAllocatedToBooking(memberName, bookingID) Then
        MsgBox memberName & " isn't allocated to booking " & bookingID & _
               " in 07_Team_Allocation -- only the crew assigned to this " & _
               "booking can check its items off.", _
               vbExclamation, "Not on this booking's crew"
        Me.Cells(r, COL_CHECKED_BY).ClearContents
        Me.Cells(r, COL_CHECKED).Value = False
        Me.Cells(r, COL_CHECKED_AT).ClearContents
    End If
End Sub

Private Sub HandleApprovedChange(r As Long)
    If Me.Cells(r, COL_APPROVED).Value = True Then
        If Me.Cells(r, COL_CHECKED).Value <> True Then
            MsgBox "This item hasn't been checked yet -- it can't be " & _
                   "approved before it's checked.", _
                   vbExclamation, "Nothing to approve"
            Me.Cells(r, COL_APPROVED).Value = False
            Exit Sub
        End If
        Me.Cells(r, COL_APPROVED_BY).Value = Environ("Username")
        Me.Cells(r, COL_APPROVED_AT).Value = Now
    Else
        Me.Cells(r, COL_APPROVED_BY).ClearContents
        Me.Cells(r, COL_APPROVED_AT).ClearContents
    End If
End Sub
