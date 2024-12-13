Function MyFunction(param1)
  If IsEmpty(param1) Then
    param1 = ""  ' Avoid error when param1 is empty
  End If
  ' ... rest of your function code ...
End Function