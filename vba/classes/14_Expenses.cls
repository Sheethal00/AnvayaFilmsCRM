Option Explicit

' Merged into the 14_Expenses worksheet's built-in code module.
' See vba/classes/02_CRM_Leads.cls header note for how this gets wired in.
'
' Column layout: A ExpenseID, B Date, C Category, D Description,
' E Amount, F LinkedBookingID, G PaidTo, H PaymentMethod.

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub

    On Error GoTo CleanUp
    IDGenerator.AutoFillID Me, Target.Row, 1, "LastExpenseNo"
CleanUp:
    Application.EnableEvents = True
End Sub
