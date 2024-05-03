Sub total_claim_projection()



'Defining Variables on the Claims Projection Tab using codename
Public IDs()

Dim Gender As Range, age As Range, condition As Range

Set Gender = ClaimsP.Range("Gender") 'C4
Set age = ClaimsP.Range("Age") 'C5
Set condition = ClaimsP.Range("Condition") 'C6

End Sub




'Define Table we are updating
Dim PD As Range
Set PD = PHdata.Range("A1").CurrentRegion

'Creating dynamic For Loop
'Originally For i = 2 To 101


For i = 2 To PD.Rows.Count

'   Gender, Age - Bringing these datapoints into our projection tool.
Gender.Value = PD.Cells(i, "B").Value
age.Value = PD.Cells(i, "C").Value

'-----------------------------------------------------------------------------
'   Condition - If statement allows us to attribute our health condition to the values 1 and 2.
    
    If PD.Cells(i, "D").Value = 1 Then
        condition.Value = "Unhealthy"
    Else
        condition.Value = "Healthy"

    End If
    
    'This brings our total back to the Policyholder Data sheet.
   PD.Cells(i, "E").Value = ClaimsP.Cells(8, "F").Value

Next i

End Sub



Sub ResetValues()
'
' ResetValues Macro
' Clears values from Tool and Totals
'

' Clears total values from the Policyholder Data Sheet
    PHdata.Range("E2:E101").Select
    Selection.ClearContents
    
'Clears inputs from our projection tool.
    ClaimsP.Range("C4:C6").Select
    Selection.ClearContents
    
'Navigates back to Policyholder Data sheet where the macro started.
    PHdata.Range("A1").Select
End Sub
