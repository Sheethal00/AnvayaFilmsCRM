Attribute VB_Name = "IDGenerator"
Option Explicit

' Single source of ID generation for every PK in the workbook (SPEC.md §2).
' Never use =MAX()+1 in a sheet formula — sorting/filtering breaks that
' pattern and causes collisions. Counters live in 19_Settings (schema/
' 19_settings.yaml id_counters), keyed by the same `counter_key` names
' used there so build_workbook.py and this module stay in sync.

Private Const SETTINGS_SHEET As String = "19_Settings"

' Returns the next ID for the given counter key, formatted per its
' pattern (e.g. "LD-####" -> "LD-0001", "RCT-YYYY-####" -> "RCT-2026-0001"),
' and persists the incremented counter back to 19_Settings.
Public Function GenerateID(counterKey As String) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)

    Dim cell As Range
    Set cell = FindCounterCell(ws, counterKey)
    If cell Is Nothing Then
        Err.Raise vbObjectError + 1, "IDGenerator.GenerateID", _
            "Unknown ID counter key: " & counterKey & _
            ". Add it to schema/19_settings.yaml id_counters and rebuild."
    End If

    Dim pattern As String
    pattern = cell.Offset(0, 1).Value ' pattern column, e.g. "LD-####"

    Dim resetsYearly As Boolean
    resetsYearly = (InStr(pattern, "YYYY") > 0)

    Dim lastNum As Long
    Dim lastYear As Long
    ParseCounterValue cell.Offset(0, 2).Value, lastNum, lastYear

    Dim nextNum As Long
    Dim currentYear As Long
    currentYear = Year(Date)

    If resetsYearly And lastYear <> currentYear Then
        nextNum = 1
    Else
        nextNum = lastNum + 1
    End If

    GenerateID = FormatID(pattern, nextNum, currentYear)

    ' Persist "num" or "year:num" back into the counter's value cell.
    If resetsYearly Then
        cell.Offset(0, 2).Value = currentYear & ":" & nextNum
    Else
        cell.Offset(0, 2).Value = nextNum
    End If
End Function

' Called from a sheet's Worksheet_Change (SPEC.md §6: IDs generated "on
' new row save"). A row counts as "new" once any other cell in it has
' content and its ID cell is still blank — avoids firing on header edits
' or on blank rows created by deleting content.
Public Sub AutoFillID(ws As Worksheet, changedRow As Long, idColumn As Long, counterKey As String)
    If changedRow < 2 Then Exit Sub ' header row

    Dim idCell As Range
    Set idCell = ws.Cells(changedRow, idColumn)
    If idCell.Value <> "" Then Exit Sub ' already has an ID

    If Not RowHasContent(ws, changedRow, idColumn) Then Exit Sub

    On Error GoTo CleanUp
    Application.EnableEvents = False
    idCell.Value = GenerateID(counterKey)
CleanUp:
    Application.EnableEvents = True
End Sub

Private Function RowHasContent(ws As Worksheet, r As Long, excludeColumn As Long) As Boolean
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To lastCol
        If c <> excludeColumn And ws.Cells(r, c).Value <> "" Then
            RowHasContent = True
            Exit Function
        End If
    Next c
    RowHasContent = False
End Function

Private Function FindCounterCell(ws As Worksheet, counterKey As String) As Range
    Dim found As Range
    Set found = ws.Columns("A:A").Find(What:=counterKey, LookAt:=xlWhole)
    Set FindCounterCell = found
End Function

Private Sub ParseCounterValue(raw As Variant, ByRef num As Long, ByRef yr As Long)
    Dim s As String
    s = CStr(raw)
    If s = "" Then
        num = 0
        yr = 0
    ElseIf InStr(s, ":") > 0 Then
        Dim parts() As String
        parts = Split(s, ":")
        yr = CLng(parts(0))
        num = CLng(parts(1))
    Else
        num = CLng(s)
        yr = 0
    End If
End Sub

Private Function FormatID(pattern As String, num As Long, yr As Long) As String
    Dim hashCount As Long
    hashCount = Len(pattern) - Len(Replace(pattern, "#", ""))

    Dim padded As String
    padded = Format(num, String(hashCount, "0"))

    Dim result As String
    result = pattern
    result = Replace(result, "YYYY", CStr(yr))
    ' Replace the run of #'s with the zero-padded number.
    Dim hashRun As String
    hashRun = String(hashCount, "#")
    result = Replace(result, hashRun, padded)

    FormatID = result
End Function
