Attribute VB_Name = "CommonUtils"
Option Explicit
'Description: collection of functions to be used in various VBA projects

'@Method Incr - incriments number by one and return it
Public Function Incr(ByRef num As Long) As Long
  num = num + 1
  Incr = num
End Function

'@Method Decr - decriments number by one and return it
Public Function Decr(ByRef num As Long) As Long
  num = num - 1
  Decr = num
End Function

'Method ToBool - force convertion to Boolean
Public Function ToBool(ByRef val As Variant) As Boolean
  On Error GoTo CatchError
  Select Case VarType(val)
    Case vbNull
      ToBool = Not IsNull(val)
    Case vbString
      ToBool = (val <> vbNullString)
    Case vbObject
      ToBool = Not val Is Nothing
    Case Is > vbArray
      On Error Resume Next
      ToBool = UBound(val) > 0
      If Err.Number = 9 Then ToBool = False
    Case vbBoolean
      ToBool = val
    Case Else
      ToBool = CBool(val)
  End Select
  Exit Function
CatchError:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

'Method StringFormat - replace placeholders in input text with parameters values
Function StringFormat(ByVal text As String, ParamArray vals() As Variant)
  Dim str As Variant, i As Long
  On Error GoTo CatchError
  For Each str In vals
    i = i + 1
    text = Replace(text, "%" & i, str)
  Next str
  StringFormat = text
CatchError:
  If Err.Number <> 0 Then Err.Raise Err.Number, "Utils.StringFormat", Err.Description
End Function
