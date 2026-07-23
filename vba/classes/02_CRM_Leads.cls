Option Explicit

' Merged into the 02_CRM_Leads worksheet's built-in code module by
' build_workbook.py (matched by filename stem to sheet name) -- not
' imported as a new component. See scripts/build_workbook.py try_import_vba.
'
' Column layout: A LeadID, ... J Status, ... O ConvertedClientID.

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Bulk paste (multi-cell/multi-row) is not handled -- fill rows one
    ' at a time, or run a cleanup macro afterward. Known limitation.
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub

    On Error GoTo CleanUp

    IDGenerator.AutoFillID Me, Target.Row, 1, "LastLeadID"

    If Target.Column = 10 Then ' Status
        If Me.Cells(Target.Row, 10).Value = "Won" Then
            LeadConversion.ConvertLeadToClient Target.Row
        End If
    End If

CleanUp:
    Application.EnableEvents = True
End Sub
