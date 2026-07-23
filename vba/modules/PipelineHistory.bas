Attribute VB_Name = "PipelineHistory"
Option Explicit

' SPEC.md §3 "10_Production_Pipeline" automation: log every stage change
' to 10a_Pipeline_History so per-editor turnaround-time reporting
' (§18 "Employee Performance") has real data to draw from, instead of
' only ever seeing the current stage.

Private Const PIPELINE_SHEET As String = "10_Production_Pipeline"
Private Const HISTORY_SHEET As String = "10a_Pipeline_History"

' Column layout (schema/10_production_pipeline.yaml): A BookingID,
' B CurrentStage, C StageEnteredDate, D AssignedEditor,
' E DaysInCurrentStage, F Notes.
Private Const COL_BOOKING_ID As Long = 1
Private Const COL_CURRENT_STAGE As Long = 2
Private Const COL_STAGE_ENTERED_DATE As Long = 3

' Column layout (schema/10a_pipeline_history.yaml): A BookingID,
' B Stage, C EnteredDate, D ExitedDate.

Public Sub LogStageChange(pipelineRow As Long)
    Dim wsPipeline As Worksheet, wsHistory As Worksheet
    Set wsPipeline = ThisWorkbook.Worksheets(PIPELINE_SHEET)
    Set wsHistory = ThisWorkbook.Worksheets(HISTORY_SHEET)

    Dim bookingID As String, newStage As String
    bookingID = wsPipeline.Cells(pipelineRow, COL_BOOKING_ID).Value
    newStage = wsPipeline.Cells(pipelineRow, COL_CURRENT_STAGE).Value
    If bookingID = "" Or newStage = "" Then Exit Sub

    On Error GoTo CleanUp
    Application.EnableEvents = False

    CloseOpenHistoryRow wsHistory, bookingID

    Dim newHistRow As Long
    newHistRow = wsHistory.Cells(wsHistory.Rows.Count, "A").End(xlUp).Row + 1
    wsHistory.Cells(newHistRow, 1).Value = bookingID
    wsHistory.Cells(newHistRow, 2).Value = newStage
    wsHistory.Cells(newHistRow, 3).Value = Date
    ' ExitedDate (column 4) stays blank -- this is now the open stage.

    wsPipeline.Cells(pipelineRow, COL_STAGE_ENTERED_DATE).Value = Date

CleanUp:
    Application.EnableEvents = True
End Sub

' At most one history row per booking has a blank ExitedDate at a time
' (the currently-open stage). Close it out before opening the next one.
Private Sub CloseOpenHistoryRow(wsHistory As Worksheet, bookingID As String)
    Dim lastRow As Long
    lastRow = wsHistory.Cells(wsHistory.Rows.Count, "A").End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        If wsHistory.Cells(r, 1).Value = bookingID And wsHistory.Cells(r, 4).Value = "" Then
            wsHistory.Cells(r, 4).Value = Date
            Exit Sub
        End If
    Next r
End Sub
