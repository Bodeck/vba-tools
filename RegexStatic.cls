VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegexStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Static class (no need to initiate) wrapper for VBS Regular Expressions"
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@ModuleDescription "Static class (no need to initiate) wrapper for VBS Regular Expressions"
'@Folder("vba-tools.Regex")
Option Explicit
'@PredeclaredId
#Const LateBinding = False

Private Type PrivateFields
  #If LateBindig Then
    Regex As Object
  #Else
    Regex As RegExp
  #End If
    Name As String
    Global As Boolean
    IgnoreCase As Boolean
    Multiline As Boolean
End Type

Private this As PrivateFields

Private Sub Class_Initialize()
  With this
  #If LateBinding Then
    Set .Regex = CreateObject("vbscript.regexp")
  #Else
    Set .Regex = New RegExp
  #End If
    'SET DEFAULTS
    .Global = True
    .IgnoreCase = True
    .Multiline = True
  End With
End Sub

Private Sub Class_Terminate()
  Set this.Regex = Nothing
End Sub

Public Property Get Name() As String
  Name = TypeName(this.Regex)
End Property

Public Property Get IsGlobal() As Boolean
  IsGlobal = this.Global
End Property

Public Property Let IsGlobal(ByVal vNewValue As Boolean)
  this.Global = vNewValue
End Property

Public Property Get IgnoreCase() As Boolean
  IgnoreCase = this.IgnoreCase
End Property

Public Property Let IgnoreCase(ByVal vNewValue As Boolean)
  this.IgnoreCase = vNewValue
End Property

Public Property Get Multiline() As Boolean
  Multiline = this.Multiline
End Property

Public Property Let Multiline(ByVal vNewValue As Boolean)
  this.Multiline = vNewValue
End Property

Public Sub QuickSetUp(ByVal IsGlobal As Boolean, ByVal Multiline As Boolean, ByVal IgnoreCase As Boolean)
  With this
    .Global = IsGlobal
    .IgnoreCase = IgnoreCase
    .Multiline = Multiline
  End With
End Sub

Public Function GetMatch(ByRef srcString As String, ByRef matchPattern As String) As String
  Dim matchObj As Object
  Dim matchColl As Object
  Dim matchStr As String
  With this.Regex
    .Global = this.Global
    .Multiline = this.Multiline
    .pattern = matchPattern
  End With
  
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
  
  With this
    .Regex.Global = True 'Ignore class property to force global to get all matches
    .Regex.IgnoreCase = this.IgnoreCase
    .Regex.Multiline = this.Multiline
    .Regex.pattern = matchPattern
  End With
  
  Set GetAllMatches = this.Regex.Execute(srcString)

End Function

'@Description ("GetNthMatch - returns match/submatch value from match collection by zero base match/submatch index.
'If there is no match/submatch for specific index returns vbNullString")
Public Function GetNthMatch(ByRef srcString As String, ByRef matchPattern As String, _
                            Optional ByVal matchIndex As Integer, Optional ByVal submatchIndex As Integer = -1) As String
  Dim matchColl As Object
  Set matchColl = GetAllMatches(srcString, matchPattern)
  
  If matchColl.Count > matchIndex Then
    Dim matchObj As Object
    Set matchObj = matchColl.Item(matchIndex)
    If submatchIndex = -1 And matchObj.SubMatches.Count = 0 Then
      Dim matchStr As String
      matchStr = matchObj.Value
    ElseIf submatchIndex < 0 And matchObj.SubMatches.Count > 0 Then
      matchStr = matchObj.SubMatches(0)
    ElseIf matchObj.SubMatches.Count > submatchIndex Then
      matchStr = matchObj.SubMatches(submatchIndex)
    End If
  End If
  GetNthMatch = matchStr

End Function

Public Function TestMatch(ByRef srcString As String, ByRef testPattern As String) As Boolean
  
  this.Regex.Global = this.Global
  this.Regex.Multiline = this.Multiline
  this.Regex.pattern = testPattern
  
  TestMatch = this.Regex.Test(srcString)
End Function

Public Function ReplaceStr(ByRef srcString As String, ByRef matchString As String, ByVal replaceWith As String)
  With this
    .Regex.Global = .Global
    .Regex.pattern = matchString
    .Regex.IgnoreCase = .IgnoreCase
    .Regex.Multiline = .Multiline
    srcString = .Regex.Replace(srcString, replaceWith)
  End With
  ReplaceStr = srcString
End Function


