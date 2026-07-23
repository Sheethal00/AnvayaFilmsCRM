Attribute VB_Name = "AvailabilityCheck"
Option Explicit

' Double-booking checks for 07_Team_Allocation and 08_Equipment
' (SPEC.md §3 "07_Team_Allocation" automation, §9 validation rules).
' Called from Worksheet_Change on the respective sheet before a new/edited
' row is accepted; returns True if the row is safe to save.

Private Const ALLOCATION_SHEET As String = "07_Team_Allocation"
Private Const EQUIPMENT_SHEET As String = "08_Equipment"
Private Const BOOKINGS_SHEET As String = "05_Bookings"

' A team member cannot be allocated to two Bookings on the same EventDate.
Public Function CheckTeamAvailability(teamMemberID As String, _
                                       eventDate As Date, _
                                       bookingID As String) As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(ALLOCATION_SHEET)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row ' BookingID column

    Dim conflictCount As Long
    conflictCount = 0

    Dim i As Long
    For i = 2 To lastRow ' row 1 = headers
        If ws.Cells(i, "C").Value = teamMemberID And _ ' TeamMemberID column
           ws.Cells(i, "E").Value = eventDate And _     ' EventDate column
           ws.Cells(i, "B").Value <> bookingID Then     ' different booking
            conflictCount = conflictCount + 1
        End If
    Next i

    If conflictCount > 0 Then
        MsgBox "Double-booking warning: this team member is already " & _
               "allocated to another booking on " & Format(eventDate, "dd-mmm-yyyy") & ".", _
               vbExclamation, "Availability Check"
        CheckTeamAvailability = False
    Else
        CheckTeamAvailability = True
    End If
End Function

' Equipment cannot be assigned to two Bookings with overlapping EventDates.
Public Function CheckEquipmentAvailability(equipmentID As String, _
                                            bookingID As String) As Boolean
    Dim wsEquip As Worksheet, wsBook As Worksheet
    Set wsEquip = ThisWorkbook.Worksheets(EQUIPMENT_SHEET)
    Set wsBook = ThisWorkbook.Worksheets(BOOKINGS_SHEET)

    Dim eqRow As Range
    Set eqRow = wsEquip.Columns("A:A").Find(What:=equipmentID, LookAt:=xlWhole)

    If eqRow Is Nothing Then
        CheckEquipmentAvailability = True ' unknown equipment ID, let higher-level validation catch it
        Exit Function
    End If

    Dim currentBookingID As String
    currentBookingID = eqRow.Offset(0, 5).Value ' CurrentBookingID column
    Dim status As String
    status = eqRow.Offset(0, 4).Value ' Status column

    If status = "Assigned" And currentBookingID <> "" And currentBookingID <> bookingID Then
        ' Only a true conflict if the two bookings' EventDates overlap —
        ' look up both WeddingDates and compare.
        If BookingDatesOverlap(wsBook, currentBookingID, bookingID) Then
            MsgBox "Equipment " & equipmentID & " is already assigned to " & _
                   "booking " & currentBookingID & " on an overlapping date.", _
                   vbExclamation, "Availability Check"
            CheckEquipmentAvailability = False
            Exit Function
        End If
    End If

    CheckEquipmentAvailability = True
End Function

Private Function BookingDatesOverlap(wsBook As Worksheet, bookingA As String, bookingB As String) As Boolean
    Dim dateA As Variant, dateB As Variant
    dateA = LookupWeddingDate(wsBook, bookingA)
    dateB = LookupWeddingDate(wsBook, bookingB)

    If IsEmpty(dateA) Or IsEmpty(dateB) Then
        BookingDatesOverlap = False
    Else
        BookingDatesOverlap = (CDate(dateA) = CDate(dateB))
    End If
End Function

Private Function LookupWeddingDate(wsBook As Worksheet, bookingID As String) As Variant
    Dim found As Range
    Set found = wsBook.Columns("A:A").Find(What:=bookingID, LookAt:=xlWhole)
    If found Is Nothing Then
        LookupWeddingDate = Empty
    Else
        LookupWeddingDate = found.Offset(0, 7).Value ' WeddingDate column
    End If
End Function
