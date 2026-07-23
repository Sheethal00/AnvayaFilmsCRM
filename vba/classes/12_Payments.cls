Option Explicit

' Merged into the 12_Payments worksheet's built-in code module.
' See vba/classes/02_CRM_Leads.cls header note for how this gets wired in.
'
' Column layout: A PaymentID, B BookingID, C ReceiptNumber, D Amount,
' E PaymentDate, F Method, G PaymentType, H Status.
'
' Two independent counters share this sheet: PaymentID never resets,
' ReceiptNumber resets every calendar year (schema/19_settings.yaml).

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub

    On Error GoTo CleanUp
    IDGenerator.AutoFillID Me, Target.Row, 1, "LastPaymentID"
    IDGenerator.AutoFillID Me, Target.Row, 3, "LastReceiptNo"
CleanUp:
    Application.EnableEvents = True
End Sub
