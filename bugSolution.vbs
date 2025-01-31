Function MyFunction(param1)
  If VarType(param1) = vbEmpty Then
    WScript.Echo "Parameter is empty"
    Exit Function
  ElseIf Len(param1) = 0 Then
    WScript.Echo "Parameter is an empty string"
    Exit Function
  End If
  ' ... rest of the function ...
End Function

MyFunction