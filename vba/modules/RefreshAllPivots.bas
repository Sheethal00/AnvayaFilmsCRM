Attribute VB_Name = "RefreshAllPivots"
Option Explicit

' Refreshes every PivotTable/PivotCache in the workbook. Bound to
' Workbook_Open and to a manual "Refresh Dashboard" button on 01_Dashboard
' (SPEC.md §3, §6). Dashboard cells are formulas/pivot references, not
' volatile formulas, so this explicit refresh step is what keeps them
' current without slowing the workbook down on every keystroke.

Public Sub RefreshAllPivots()
    Dim pc As PivotCache
    For Each pc In ThisWorkbook.PivotCaches
        pc.Refresh
    Next pc

    Application.CalculateFull
End Sub
