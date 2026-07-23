Option Explicit

' Merged into the 15_Freelancers worksheet's built-in code module.
' See vba/classes/02_CRM_Leads.cls header note for how this gets wired in.
'
' Column layout: A FreelancerID, B Name, C Role, D Phone, E RatePerEvent,
' F BankDetails, G TotalPaidYTD.

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub

    On Error GoTo CleanUp
    IDGenerator.AutoFillID Me, Target.Row, 1, "LastFreelancerID"
CleanUp:
    Application.EnableEvents = True
End Sub
