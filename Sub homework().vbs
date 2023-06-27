Sub homework()

' declaration of variables

Dim ticker As String

Dim vol As Double

Dim opencell As Double

Dim closecell As Double

Dim position As Double

Dim Lastrow As Long

Dim increase As Double

Dim decrease As Double

Dim maxvol As Double

Dim increaseticker As String
Dim decreaseticker As String
Dim maxvolticker As String





' initiation of variables

position = 2

vol = 0

opencell = Cells(2, 3)

ticker = Cells(2, 1)

closecell = 0

' last row determination

Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Cells(1, 9).Value = ("ticker")

Cells(1, 10).Value = ("Yearly change")

Cells(1, 11).Value = ("Percent Change")

Cells(1, 12).Value = ("Total stock volume")


' looping through all rows of ticker collumn

For i = 2 To Lastrow

'condition for same ticker

    If Cells(i, 1).Value = Cells(i + 1, 1).Value Then

'count sum for volume


    vol = vol + Cells(i, 7).Value

    Else
    
    vol = vol + Cells(i, 7).Value
    
    'affectation

    Cells(position, 9).Value = ticker
    
    closecell = Cells(i, 6).Value
    
    Cells(position, 10).Value = closecell - opencell
    
    Cells(position, 11).Value = ((closecell - opencell) / opencell) * 100
    
    Cells(position, 12).Value = vol
    
    'reinitiation
    
    vol = 0
    
   opencell = Cells(i + 1, 3).Value
    
    position = position + 1
    
    ticker = Cells(i + 1, 1).Value
    
    End If
    
    Next
    
    

increase = Cells(2, 10).Value

decrease = Cells(2, 10).Value

maxvol = Cells(2, 11).Value
    
    
    
    
    For i = 2 To Lastrow
    
    ' Greates increase determination
    
    If Cells(i, 11).Value > increase Then
    
    increase = Cells(i, 11).Value
    
    increaseticker = Cells(i, 9).Value
    
    End If
    
    If Cells(i, 11).Value < decrease Then
    
    decrease = Cells(i, 11).Value
    
    decreaseticker = Cells(i, 9).Value
    
    End If
    
    If Cells(i, 12).Value > maxvol Then
    
    maxvol = Cells(i, 12).Value
    
    maxvolticker = Cells(i, 9).Value
    
    End If
    
Next

Cells(1, 16).Value = ("ticker")

Cells(1, 17).Value = ("Value")

Cells(2, 15).Value = ("Greatest % increase")

Cells(3, 15).Value = ("Greatest % decrease")

Cells(4, 15).Value = ("Greatest Total Volume")

Cells(2, 16).Value = increaseticker

Cells(2, 17).Value = increase

Cells(3, 16).Value = decreaseticker

Cells(3, 17).Value = decrease

Cells(4, 16).Value = maxvolticker

Cells(4, 17).Value = maxvol

For i = 2 To Lastrow

If Cells(i, 10).Value < 0 Then

Cells(i, 10).Interior.ColorIndex = 3

Else

Cells(i, 10).Interior.ColorIndex = 4

End If

Next


End Sub
