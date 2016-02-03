VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VB_ExCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'require:
'Ref:Visual Basic For Applications
Option Explicit

Private mKeyValueCol As VBA.Collection
Private mKeyCol As VBA.Collection
'【私有-函数】
Private Function exists(ByVal key As Variant) As Boolean
  On Error GoTo errHandle
  mKeyValueCol.Item key
  exists = True
  Exit Function
errHandle:
  exists = False
End Function


'【公有-属性】
Public Property Get Item(key As Variant) As Variant
  If exists(key) Then
    Item = mKeyValueCol.Item(key)
  Else
    Set Item = Nothing
  End If
End Property

Public Property Get Count() As Long
  Count = mKeyValueCol.Count
End Property

'【公开-方法】
Public Sub SetItem(key As Variant, value As Variant)
  If exists(key) Then
    mKeyValueCol.Remove key
    mKeyCol.Remove key
  End If
  mKeyValueCol.Add value, key
  mKeyCol.Add key, key
End Sub

Public Function SafeGetItem(key As Variant) As Variant
  If exists(key) Then
    If VarType(mKeyValueCol.Item(key)) = vbObject Then
      Set SafeGetItem = mKeyValueCol.Item(key)
    Else
      SafeGetItem = mKeyValueCol.Item(key)
    End If
  Else
    SafeGetItem = ""
  End If
End Function

Public Sub Remove(key As Variant)
  If exists(key) Then
    mKeyValueCol.Remove key
    mKeyCol.Remove key
  End If
End Sub

Public Function AllKeys() As String()
  Dim pKeys() As String
  Dim I As Integer
  Dim keyCount As Integer
  
  keyCount = mKeyCol.Count
  
  If keyCount <= 0 Then
    Err.Raise 10001, , "不存在数据"
    Exit Function
  End If
  
  ReDim pKeys(mKeyCol.Count - 1)
  
  For I = 0 To mKeyCol.Count - 1
    pKeys(I) = mKeyCol.Item(I + 1)
  Next I
  AllKeys = pKeys
End Function

Public Function JoinKeys(Optional ByVal Seperator As String = "") As String
  Dim keys() As String
  Dim concatStr As String
  Dim I As Integer
  
  If mKeyCol.Count <= 0 Then
    JoinKeys = ""
    Exit Function
  End If
  
  keys = AllKeys()
  
  For I = 0 To UBound(keys)
    If concatStr = "" Then
      concatStr = keys(I)
    Else
      concatStr = concatStr & Seperator & keys(I)
    End If
  Next I
  JoinKeys = concatStr
End Function

Public Function Join(Optional ByVal Seperator As String = "") As String
  Dim joinStr As String
  Dim I As Integer
  Dim key As String
  
  If mKeyCol.Count <= 0 Then
    Join = ""
  End If
  
  joinStr = ""
  
  For I = 1 To mKeyCol.Count
    key = mKeyCol.Item(I)
    If I = 1 Then
      joinStr = joinStr & key & "=" & mKeyValueCol.Item(key)
    Else
      joinStr = joinStr & Seperator & key & "=" & mKeyValueCol.Item(key)
    End If
  Next I
  
  Join = joinStr
End Function

Public Function ToJson() As String
  Dim jsonStr As String
  Dim I As Integer
  Dim key As String
  
  If mKeyCol.Count <= 0 Then
    ToJson = "[]"
  End If
  
  jsonStr = "["
  
  For I = 1 To mKeyCol.Count
    key = mKeyCol.Item(I)
    If I = 1 Then
      jsonStr = jsonStr & "{""" & Replace(key, """", "\""") & """:""" & Replace(mKeyValueCol.Item(key), """", "\""") & """}"
    Else
      jsonStr = jsonStr & ",{""" & Replace(key, """", "\""") & """:""" & Replace(mKeyValueCol.Item(key), """", "\""") & """}"
    End If
  Next I
  
  jsonStr = jsonStr & "]"
  ToJson = jsonStr
End Function


'类初始化
Private Sub Class_Initialize()
  Set mKeyValueCol = New VBA.Collection
  Set mKeyCol = New VBA.Collection
End Sub

Private Sub Class_Terminate()
  Set mKeyValueCol = Nothing
  Set mKeyCol = Nothing
End Sub


