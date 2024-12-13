Function MyFunction(param1)
  If IsEmpty(param1) Then
    param1 = ""  ' Provide a default value if parameter is empty
  End If
  ' ... rest of your function code ...
  If IsEmpty(param1) Then
    result = "Parameter is empty"
  Else
    result = "Parameter value: " & param1
  End If
  MyFunction = result
End Function

' Example usage
MsgBox MyFunction(10)  ' Output: Parameter value: 10
MsgBox MyFunction("")   ' Output: Parameter is empty
MsgBox MyFunction()     ' Output: Parameter is empty