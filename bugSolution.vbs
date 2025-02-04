Function MyFunction(param1)
  On Error Resume Next
  If IsEmpty(param1) Then
    Err.Raise 9999, , "Parameter cannot be empty"
  End If
  On Error GoTo 0
  If Err.Number <> 0 Then
    'Handle custom error 9999
    MsgBox "Error: " & Err.Description
    Err.Clear
  End If
  ' ... rest of the function
End Function