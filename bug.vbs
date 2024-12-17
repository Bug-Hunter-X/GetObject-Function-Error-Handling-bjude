Function GetObject(progID)
  On Error Resume Next
  Set obj = GetObject(progID)
  If Err.Number <> 0 Then
    Set obj = CreateObject(progID)
  End If
  Set GetObject = obj
End Function

'This will cause an error if the object is not found:
Set myObj = GetObject("Some.Unknown.Object")