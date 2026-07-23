VERSION 1.0 CLASS
BEGIN
  MultiUseModule = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Workbook-level triggers (SPEC.md §3, §4, §6):
'   - Refresh dashboard pivots on open
'   - Check payment reminders on open
'   - Hide/unhide sheet tabs per Windows username (UI-level only —
'     CLAUDE.md caveat: NOT real access control, anyone with the VBA
'     password or a copy of the file can see everything)

Private Sub Workbook_Open()
    On Error Resume Next ' keep the workbook usable even if one step fails

    RefreshAllPivots.RefreshAllPivots
    PaymentReminder.CheckDuePayments
    ApplyRolePermissions

    On Error GoTo 0
End Sub

Private Sub ApplyRolePermissions()
    Dim currentUser As String
    currentUser = Environ("Username")

    Dim wsSettings As Worksheet
    Set wsSettings = ThisWorkbook.Worksheets("19_Settings")

    Dim userRow As Range
    Set userRow = FindUserRow(wsSettings, currentUser)

    Dim role As String
    If userRow Is Nothing Then
        role = "" ' unknown user — leave all sheets as-is rather than lock the file
        Exit Sub
    Else
        role = userRow.Offset(0, 1).Value ' Role column
    End If

    ShowSheetsForRole role
End Sub

Private Function FindUserRow(ws As Worksheet, username As String) As Range
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects("UserRolePermissions")
    On Error GoTo 0

    If tbl Is Nothing Then Exit Function

    Set FindUserRow = tbl.ListColumns("Username").DataBodyRange.Find( _
        What:=username, LookAt:=xlWhole)
End Function

Private Sub ShowSheetsForRole(role As String)
    Dim ws As Worksheet

    ' Admin sees everything — reset visibility and exit.
    If role = "Admin" Or role = "" Then
        For Each ws In ThisWorkbook.Worksheets
            ws.Visible = xlSheetVisible
        Next ws
        Exit Sub
    End If

    Dim allowedPrefixes As Variant
    Select Case role
        Case "Manager"
            allowedPrefixes = Array("01", "02", "03", "04", "05", "06", "07", _
                                     "12", "13", "14", "15", "16", "17", "18")
        Case "Editor"
            allowedPrefixes = Array("09", "10", "11")
        Case "Sales"
            allowedPrefixes = Array("01", "02", "03", "04")
        Case "Accountant"
            allowedPrefixes = Array("12", "13", "14", "15", "16", "17", "18")
        Case Else
            allowedPrefixes = Array() ' unrecognized role -> show nothing but Dashboard
    End Select

    For Each ws In ThisWorkbook.Worksheets
        If SheetPrefixAllowed(ws.Name, allowedPrefixes) Then
            ws.Visible = xlSheetVisible
        Else
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
End Sub

Private Function SheetPrefixAllowed(sheetName As String, allowedPrefixes As Variant) As Boolean
    Dim prefix As Variant
    For Each prefix In allowedPrefixes
        If Left(sheetName, Len(CStr(prefix))) = CStr(prefix) Then
            SheetPrefixAllowed = True
            Exit Function
        End If
    Next prefix
    SheetPrefixAllowed = False
End Function
