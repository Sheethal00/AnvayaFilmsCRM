Option Explicit

' Merged into the 05_Bookings worksheet's built-in code module.
' See vba/classes/02_CRM_Leads.cls header note for how this gets wired in.
'
' Column layout: A BookingID, B ClientID, C QuoteID, D EventType,
' E PackageValue, F AdvanceAmount, G AdvanceDate, H BalanceAmount,
' I WeddingDate, J VenueName, K VenueMapLink, L DeliveryDeadline,
' M Status.

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub

    On Error GoTo CleanUp
    IDGenerator.AutoFillID Me, Target.Row, 1, "LastBookingID"
CleanUp:
    Application.EnableEvents = True
End Sub
