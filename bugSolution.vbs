Function MyFunction(param)
  'More robust handling of empty or Null parameters
  If IsEmpty(param) Or param = "" Or param Is Nothing Then
    'Handle empty or Null param
  Else
    'Process the parameter
  End If
End Function