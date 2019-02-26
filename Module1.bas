Attribute VB_Name = "Module1"
Sub stockcalc()
  Dim tickername As String
  Dim tickertotal As Double
  Dim activerow As Integer
  Dim lastrow As Double
  tickertotal = 0
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  MsgBox (lastrow)
  activerow = 2
  For i = 2 To lastrow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      tickername = Cells(i, 1).Value
      tickertotal = tickertotal + Cells(i, 7).Value
      Range("i" & activerow).Value = tickername
      Range("l" & activerow).Value = tickertotal
      activerow = activerow + 1
      tickertotal = 0
      
    Else
      tickertotal = tickertotal + Cells(i, 7).Value
      Range("l" & activerow).Value = tickertotal
    End If
  Next i
  Dim openingvalue As Double
  Dim closingvalue As Double
  activerow = 2
  For i = 2 To lastrow
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
      openingvalue = Cells(i, 3).Value
      closingvalue = Cells(i + 260, 3).Value
      Range("j" & activerow).Value = closingvalue - openingvalue
      If Range("j" & activerow).Value > 0 Then
        Range("j" & activerow).Interior.ColorIndex = 4
      ElseIf Range("j" & activerow).Value < 0 Then
        Range("j" & activerow).Interior.ColorIndex = 3
      End If
        Range("k" & activerow).Value = (((closingvalue - openingvalue) / (openingvalue + 0.00001)))
      activerow = activerow + 1
    End If
  Next i
End Sub

