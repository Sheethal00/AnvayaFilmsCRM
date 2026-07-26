Attribute VB_Name = "ChecklistWorkflow"
Option Explicit

' Crew-involvement check for 09_Shoot_Checklist.CheckedBy: confirms the
' named person is actually allocated to the booking in
' 07_Team_Allocation, not just any team member ("the pre check has to be
' done by the member who is involved").

Private Const TEAM_LIST_SHEET As String = "19_Settings"
Private Const ALLOCATION_SHEET As String = "07_Team_Allocation"
Private Const TEAM_LIST_NAME_COL As String = "I" ' 19_Settings team_list: H TeamMemberID, I Name
Private Const TEAM_LIST_ID_COL As String = "H"

Public Function IsAllocatedToBooking(memberName As String, bookingID As String) As Boolean
    Dim teamMemberID As String
    teamMemberID = LookupTeamMemberID(memberName)
    If teamMemberID = "" Then
        IsAllocatedToBooking = False
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(ALLOCATION_SHEET)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row ' BookingID column

    Dim i As Long
    For i = 2 To lastRow ' row 1 = headers
        If ws.Cells(i, "B").Value = bookingID And ws.Cells(i, "C").Value = teamMemberID Then
            IsAllocatedToBooking = True
            Exit Function
        End If
    Next i

    IsAllocatedToBooking = False
End Function

Private Function LookupTeamMemberID(memberName As String) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(TEAM_LIST_SHEET)

    Dim found As Range
    Set found = ws.Columns(TEAM_LIST_NAME_COL & ":" & TEAM_LIST_NAME_COL).Find( _
        What:=memberName, LookAt:=xlWhole)

    If found Is Nothing Then
        LookupTeamMemberID = ""
    Else
        LookupTeamMemberID = ws.Cells(found.Row, TEAM_LIST_ID_COL).Value
    End If
End Function
