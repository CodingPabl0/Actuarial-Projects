Sub total_claim_projection()

For i = 2 To 101

'   Gender, Age - Bringing these datapoints into our projection tool.
Sheets("Claims Projection Tool").Cells(4, "C").Value = Sheets("Policyholder Data").Cells(i, "B").Value
Sheets("Claims Projection Tool").Cells(5, "C").Value = Sheets("Policyholder Data").Cells(i, "C").Value

'   Condition - If statement allows us to attribute our health condition to the values 1 and 2.
    If Sheets("Policyholder Data").Cells(i, "D").Value = 1 Then

        Sheets("Claims Projection Tool").Cells(6, "C").Value = "Unhealthy"

        Else
        Sheets("Claims Projection Tool").Cells(6, "C").Value = "Healthy"

    End If
    
    'This brings our total back to the Policyholder Data sheet.
   Sheets("Policyholder Data").Cells(i, "E").Value = Sheets("Claims Projection Tool").Cells(8, "F").Value

Next i

End Sub

Sub ResetValues()
'
' ResetValues Macro
' Clears values from Tool and Totals
'

' Clears total values from the Policyholder Data Sheet
    Range("E2:E101").Select
    Selection.ClearContents
    
'Clears inputs from our projection tool.
    Sheets("Claims Projection Tool").Select '
    Range("C4:C6").Select
    Selection.ClearContents
    
'Navigates back to Policyholder Data sheet where the macro started.
    Sheets("Policyholder Data").Select
    Range("A1").Select
End Sub
