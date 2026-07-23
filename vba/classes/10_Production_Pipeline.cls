Option Explicit

' Merged into the 10_Production_Pipeline worksheet's built-in code module.
' See vba/classes/02_CRM_Leads.cls header note for how this gets wired in.
'
' Column layout: A BookingID, B CurrentStage, C StageEnteredDate,
' D AssignedEditor, E DaysInCurrentStage, F Notes.
'
' No AutoFillID here: BookingID is a lookup FK the user picks from
' 05_Bookings, not a locally-generated ID.

Private Const COL_CURRENT_STAGE As Long = 2

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub
    If Target.Column <> COL_CURRENT_STAGE Then Exit Sub
    If Target.Value = "" Then Exit Sub

    PipelineHistory.LogStageChange Target.Row
End Sub
