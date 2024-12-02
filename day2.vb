Sub passorfail()

Dim i, x, y, z, pass, fail, steps, counter, inc As Integer

'inc = 0 means the step hasn't ran yet, inc = 1 is increasing and inc = 2 is decreasing

For i = 1 To 3695

inc = 0
fail = 0
pullout = 2

  For counter = 2 To 8
  
    x = Sheets("New Data").Cells(i, counter)
    y = Sheets("New Data").Cells(i, counter + 1)
    
    If counter = 2 Then
    
      If x < y Then
        inc = 1
      End If
    
      If x > y Then
        inc = 2
      End If
    
      If x = y Then
        fail = 1
        Exit For
      End If
      
    End If
    
    If y = "" Then
      counter = 8
      Exit For
    End If
    
    If inc = 1 Then
      
      z = y - x
      
      If z < 1 Then
        fail = 1
        Exit For
      End If
      
      If z > 3 Then
        fail = 1
        Exit For
      End If
      
    End If
    
    If inc = 2 Then
      
      z = x - y
      
      If z < 1 Then
        fail = 1
        Exit For
      End If
      
      If z > 3 Then
        fail = 1
        Exit For
      End If
      
    End If
    
  Next counter
  
If fail = 0 Then
  pass = pass + 1
  Sheets("New Data").Cells(i, 10) = "Pass"
End If

If fail = 1 Then
  Sheets("New Data").Cells(i, 10) = "Fail"
End If

Next i

MsgBox ("The answer is: " & pass)

End Sub

