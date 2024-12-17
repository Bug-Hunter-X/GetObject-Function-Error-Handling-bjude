Function GetObjectSafe(progID)
  On Error Resume Next
  Set obj = GetObject(progID)
  If Err.Number <> 0 Then
    Err.Clear
    'Handle the error appropriately.  Here, we're returning Nothing:
    Set obj = Nothing
  End If
  Set GetObjectSafe = obj
End Function

'This version handles errors more robustly:
Set myObj = GetObjectSafe("Some.Unknown.Object")
If myObj Is Nothing Then
  MsgBox "Object not found!", vbExclamation
Else
  'Use myObj
End If