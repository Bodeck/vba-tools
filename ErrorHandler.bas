Attribute VB_Name = "ErrorHandler"
Option Explicit

Public ErrorStack As New Collection

Public Sub RaiseError(ByVal procName As String)
  If Err.Number = 0 Then Exit Sub
  
  If ErrorStack.Count = 0 Then
    ErrorStack.Add procName
  Else
    ErrorStack.Add procName, , 1
  End If
  
  Err.Raise Err.Number, Err.Source
End Sub

Public Sub DisplayError(Optional ByRef optionDescription As String)
  Dim stringifiedErrorStack As String
  Dim errDesc As String, errSource As String
  Dim el As Variant
  
  If Err.Number = 0 Then Exit Sub
  
  Do While ErrorStack.Count > 0
    el = ErrorStack.Item(1)
    stringifiedErrorStack = Space(4) & el & vbNewLine & stringifiedErrorStack
    ErrorStack.Remove (1)
  Loop
  
  MsgBox Err.Description & vbNewLine & vbNewLine & _
          "Error Stack:" & vbNewLine & stringifiedErrorStack, vbCritical + vbOKOnly, "Error#: " & Err.Number
  
End Sub

Public Sub RemoveAll()
  Do While ErrorStack.Count > 0
    ErrorStack.Remove (1)
  Loop
End Sub

'Usage
Private Sub TopSub()
  On Error GoTo CatchError
  Call SubLevel1
  Exit Sub
CatchError:
  DisplayError "Optional description"
End Sub

Private Sub SubLevel2()
  On Error GoTo CatchError
  Dim x As Integer
  x = FunctionWithError
  Exit Sub
CatchError:
  RaiseError "ModuleName.SubLevel2"
End Sub

Private Sub SubLevel1()
  On Error GoTo CatchError
  Call SubLevel2
  Exit Sub
CatchError:
  RaiseError "ModuleName.SubLevel1"
End Sub

Private Function FunctionWithError() As Integer
  On Error GoTo CatchError
    FunctionWithError = 10 / 0
  Exit Function
CatchError:
  RaiseError "ModuleName.FunctionWithError"
End Function

