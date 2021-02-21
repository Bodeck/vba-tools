VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegexStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'@Folder: VBA-Regex
'@Description: Static class (no need to initiate) wrapper for VBS Regular Expressions

Private Type PrivateFields
  Regex As RegExp
  Name As String
End Type

Private this As PrivateFields

Public Property Get Name() As String
  Name = TypeName(this.Regex)
End Property

Private Sub Class_Initialize()
  Set this.Regex = CreateObject("vbscript.regexp")
End Sub

Public Function GetMatch(ByRef srcString As String, ByRef matchPattern As String) As String
  Dim matchObj As Object
  Dim matchColl As Object
  Dim matchStr As String
  
  this.Regex.Global = False
  this.Regex.pattern = matchPattern
  
  Set matchColl = this.Regex.Execute(srcString)
  
  If matchColl.Count > 0 Then
    Set matchObj = matchColl.Item(0)
    If matchObj.SubMatches.Count <> 0 Then
      matchStr = matchObj.SubMatches(0)
    Else
      matchStr = matchObj.Value
    End If
  End If
  
  GetMatch = matchStr

End Function

'@GetAllMatches - return Match Objects Collection
Public Function GetAllMatches(ByRef srcString As String, ByRef matchPattern As String) As Object
  Dim matchObj As Object
  Dim matchStr As String
  this.Regex.Global = True
  this.Regex.pattern = matchPattern
  Set GetAllMatches = this.Regex.Execute(srcString)

End Function

'@Method GetNthMatch - returns match/submatch value from match collection by zero base match/submatch index.
'If there is no match/submatch for specific index returns vbNullString
Public Function GetNthMatch(ByRef srcString As String, ByRef matchPattern As String, _
                            Optional ByVal matchIndex As Integer, Optional ByVal submatchIndex As Integer = -1) As String
  Dim matchObj As Object
  Dim matchColl As Object
  Dim matchStr As String
  
  Set matchColl = GetAllMatches(srcString, matchPattern)
  
  If matchColl.Count > matchIndex Then
    Set matchObj = matchColl.Item(matchIndex)
    If submatchIndex = -1 And matchObj.SubMatches.Count = 0 Then
      matchStr = matchObj.Value
    ElseIf submatchIndex < 0 And matchObj.SubMatches.Count > 0 Then
      matchStr = matchObj.SubMatches(0)
    ElseIf matchObj.SubMatches.Count > submatchIndex Then
      matchStr = matchObj.SubMatches(submatchIndex)
    End If
  End If
  GetNthMatch = matchStr

End Function

Public Function TestMatch(ByRef srcString As String, ByRef testPattern)
  this.Regex.pattern = testPattern
  TestMatch = this.Regex.Test(srcString)
End Function
