VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'请一定注意以下这一点
'说明：请在“工程”-“引用”中，选择“Microsoft VBScript Regular Expressions 5.5”
Private mRegExp     As New RegExp
Private mIgnoreCase As Boolean
Private mMultiLine  As Boolean
Private mGlobal     As Boolean


'Method
Public Sub Config(ByVal bIgnoreCase As Boolean, _
                  ByVal bMultiLine As Boolean, _
                  ByVal bGlobal As Boolean)
  mIgnoreCase = bIgnoreCase
  mMultiLine = bMultiLine
  mGlobal = bGlobal

  With mRegExp
    .IgnoreCase = mIgnoreCase
    .Global = mGlobal
    .MultiLine = mMultiLine
  End With

End Sub

Public Function GetMatchCollection(ByVal Patten As String, _
                                   ByVal Sample As String) As MatchCollection
  With mRegExp
    .Pattern = Patten
    Set GetMatchCollection = .Execute(Sample)
  End With
End Function

Public Function GetSubMatches(ByVal Patten As String, _
                              ByVal Sample As String, _
                              Optional ByVal MatchIndex As Long = 0) As SubMatches
  Dim Matches As MatchCollection

  Set Matches = GetMatchCollection(Patten, Sample)

  If Matches Is Nothing Then
    Set GetSubMatches = Nothing
  Else
    Set GetSubMatches = Matches(MatchIndex).SubMatches
  End If

End Function

Public Function GetScalarSubMatch(ByVal Patten As String, _
                                  ByVal Sample As String) As String

  Dim mSubMatches As SubMatches

  Set mSubMatches = GetSubMatches(Patten, Sample)

  '
  If mSubMatches Is Nothing Then
    GetScalarSubMatch = ""
    Exit Function
  End If

  If mSubMatches.Count > 0 Then
    GetScalarSubMatch = mSubMatches.Item(0)
  Else
    GetScalarSubMatch = ""
  End If

End Function

Public Function GetSubMatchArray(ByVal Patten As String, _
                                 ByVal Sample As String) As String()

  Dim tSubMatches As SubMatches
  Dim tArr()      As String
  Dim tCount      As Long
  Dim i           As Long

  Set tSubMatches = GetSubMatches(Patten, Sample)

  If Not tSubMatches Is Nothing Then
    tCount = tSubMatches.Count
  Else
    tCount = 0
  End If

  If tCount > 0 Then
    ReDim tArr(tCount - 1)

    For i = 0 To tCount - 1
        tArr(i) = tSubMatches(i)
    Next i

    GetSubMatchArray = tArr
  Else
    ReDim tArr(0)
    GetSubMatchArray = tArr
    Exit Function
  End If
End Function

Public Function GetAllSubMatchesArray(ByVal Patten As String, _
                                      ByVal Sample As String) As String()

  Dim tMC      As MatchCollection
  Dim tSM      As SubMatches
  Dim tC       As New Collection
  Dim tArr()   As String
  Dim tCount   As Long
  Dim i        As Long
  Dim j        As Long
  Dim tMCCount As Long
  
  Set tMC = GetMatchCollection(Patten, Sample)
    
  If tMC.Count > 0 Then

    For i = 0 To tMC.Count - 1
      Set tSM = tMC(i).SubMatches

      If tSM.Count > 0 Then
        For j = 0 To tSM.Count - 1
          tC.Add tSM(j), CStr(tCount)
          tCount = tCount + 1
        Next j
      End If

    Next i

  End If

  If tMCCount > 0 Then
    ReDim tArr(0)
  Else
    ReDim tArr(tCount - 1)

    For i = 0 To tCount - 1
      tArr(i) = tC.Item(CStr(i))
    Next i
  End If
  GetAllSubMatchesArray = tArr
End Function

