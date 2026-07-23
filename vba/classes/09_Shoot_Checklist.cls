Option Explicit

' Merged into the 09_Shoot_Checklist worksheet's built-in code module.
' See vba/classes/02_CRM_Leads.cls header note for how this gets wired in.
'
' Column layout: A BookingID, B ChecklistPhase, C ItemName, D Checked,
' E CheckedBy, F CheckedAt.

Private Const COL_CHECKED As Long = 4
Private Const COL_CHECKED_AT As Long = 6

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub
    If Target.Column <> COL_CHECKED Then Exit Sub

    On Error GoTo CleanUp
    Application.EnableEvents = False

    If Target.Value = True Then
        Me.Cells(Target.Row, COL_CHECKED_AT).Value = Now
    Else
        Me.Cells(Target.Row, COL_CHECKED_AT).ClearContents
    End If

CleanUp:
    Application.EnableEvents = True
End Sub
