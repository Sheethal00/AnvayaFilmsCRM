Attribute VB_Name = "PaymentReminder"
Option Explicit

' Daily, on-file-open check for payments due soon (SPEC.md §3, §6
' "Payment Reminder" trigger row). Flags results on 01_Dashboard rather
' than popping a modal on every open, so it doesn't get in the way of
' normal use.

Private Const PAYMENTS_SHEET As String = "12_Payments"
Private Const BOOKINGS_SHEET As String = "05_Bookings"
Private Const DASHBOARD_SHEET As String = "01_Dashboard"
Private Const REMINDER_WINDOW_DAYS As Long = 7

Public Sub CheckDuePayments()
    Dim wsBookings As Worksheet
    Set wsBookings = ThisWorkbook.Worksheets(BOOKINGS_SHEET)

    Dim lastRow As Long
    lastRow = wsBookings.Cells(wsBookings.Rows.Count, "A").End(xlUp).Row

    Dim dueList As String
    dueList = ""

    Dim i As Long
    For i = 2 To lastRow
        Dim balance As Double
        balance = wsBookings.Cells(i, "F").Value ' BalanceAmount column (formula result)

        Dim deliveryDeadline As Variant
        deliveryDeadline = wsBookings.Cells(i, "L").Value ' DeliveryDeadline column

        If balance > 0 And IsDate(deliveryDeadline) Then
            If CDate(deliveryDeadline) - Date <= REMINDER_WINDOW_DAYS And _
               CDate(deliveryDeadline) - Date >= 0 Then
                dueList = dueList & wsBookings.Cells(i, "A").Value & _
                          " (balance " & Format(balance, "#,##0") & "), "
            End If
        End If
    Next i

    WriteReminderToDashboard dueList
End Sub

Private Sub WriteReminderToDashboard(dueList As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)

    Dim reminderCell As Range
    Set reminderCell = ws.Columns("A:A").Find(What:="Payment Reminders", LookAt:=xlWhole)

    If reminderCell Is Nothing Then Exit Sub ' layout not built yet, skip silently

    If Len(dueList) > 0 Then
        reminderCell.Offset(0, 1).Value = Left(dueList, Len(dueList) - 2)
    Else
        reminderCell.Offset(0, 1).Value = "No payments due in the next " & _
            REMINDER_WINDOW_DAYS & " days."
    End If
End Sub
