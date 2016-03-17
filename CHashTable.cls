VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHashTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//////////////////////////////////////////////////////////////////////////////
'@@summary
'@@require
'@@reference
'@@license
'@@author
'@@create
'@@modify
'//////////////////////////////////////////////////////////////////////////////
'【内部采用一基数组】

'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有声明
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       接口继承
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       公有常量
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       公有数据类型
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       公有变量
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       公有API
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       事件声明
'------------------------------------------------------------------------------


'//////////////////////////////////////////////////////////////////////////////
'//
'//      私有声明
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       私有常量
'------------------------------------------------------------------------------
Private Const INIT_SIZE = 1024
Private Const INIT_GAP_SIZE = 128

'------------------------------------------------------------------------------
'       私有数据类型
'------------------------------------------------------------------------------
Private Type ArrayElement
  IsTaken As Boolean
  Value As Variant
  Key As String
End Type


'------------------------------------------------------------------------------
'       私有变量
'------------------------------------------------------------------------------
Private mCount As Long
Private mCurrentSize As Long
Private mArray() As ArrayElement
Private mSize As Long
Private mGapSize As Long


'------------------------------------------------------------------------------
'       属性变量
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       私有API
'------------------------------------------------------------------------------


'//////////////////////////////////////////////////////////////////////////////
'//
'//      类
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       初始化
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
  resizeArray INIT_SIZE
  Me.GapSize = INIT_GAP_SIZE
End Sub


'------------------------------------------------------------------------------
'       销毁
'------------------------------------------------------------------------------
Private Sub Class_Terminate()

End Sub


'//////////////////////////////////////////////////////////////////////////////
'//
'//      事件处理
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      私有属性
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      私有方法
'//
'//////////////////////////////////////////////////////////////////////////////
Private Sub resizeArray(ByVal newSize As Long)
  If mCount > 0 Then
    ReDim Preserve mArray(1 To newSize)
  Else
    ReDim mArray(1 To newSize)
  End If
  mCurrentSize = newSize
End Sub

Private Sub rearangeArray()
  '【等待施工】
  'if data exists on 1,5,9,11,....so on,which means data are ranging almost randomly,
  'we need to call this sub to rearrange it
  
End Sub

Private Sub expandArray()
  mCurrentSize = mCurrentSize + mGapSize
  resizeArray mCurrentSize
End Sub

Private Function findIndex(ByVal keyword As String) As Long
  Dim index As Long, startIndex As Long
  index = calcInitIndex(keyword)
  
  startIndex = index
  GoTo DirectExec
MoveOn:
  If startIndex = index Then
    findIndex = -1
    Exit Function
  End If
DirectExec:
  If mArray(index).IsTaken = True And mArray(index).Key = keyword Then
    findIndex = index
    Exit Function
  Else
    index = (index + 1) Mod mCurrentSize
    If index = 0 Then index = index + 1
    GoTo MoveOn
  End If
End Function

Private Function findFreeIndex(ByVal keyword As String) As Long
  Dim index As Long, startIndex As Long
  index = calcInitIndex(keyword)
  
  startIndex = index
  GoTo DirectExec
MoveOn:
  If startIndex = index Then
    findFreeIndex = -1
    Exit Function
  End If
DirectExec:
  If mArray(index).IsTaken = False Then
    findFreeIndex = index
    Exit Function
  Else
    index = (index + 1) Mod mCurrentSize
    If index = 0 Then index = index + 1
    GoTo MoveOn
  End If
End Function


Private Function calcInitIndex(ByVal keyword As String) As Long
  Dim index As Long, i As Long
  Dim bytes() As Byte
  bytes = keyword
  For i = LBound(bytes) To UBound(bytes)
    index = index + bytes(i)
    If index > mCount Then
      index = index Mod mCurrentSize
    End If
  Next i
  If index = 0 Then index = index + 1
  calcInitIndex = index
End Function

Private Function UTF8_URLEncoding(ByVal szInput As String)
  Dim wch As String, uch As String, szRet As String
  Dim x As Long
  Dim nAsc As Long, nAsc2 As Long, nAsc3 As Long
  If szInput = "" Then
    UTF8_URLEncoding = szInput
    Exit Function
  End If
  For x = 1 To Len(szInput)
    wch = Mid(szInput, x, 1)
    nAsc = AscW(wch)
    
    If nAsc < 0 Then nAsc = nAsc + 65536
    
    If (nAsc And &HFF80) = 0 Then
      szRet = szRet & wch
    Else
      If (nAsc And &HF000) = 0 Then
        uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
        szRet = szRet & uch
      Else
        uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
        Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
        Hex(nAsc And &H3F Or &H80)
        szRet = szRet & uch
      End If
    End If
  Next
  UTF8_URLEncoding = szRet
End Function

Public Function UTF8_UrlDecode(ByVal URL As String)
    Dim B, ub   ''中文字的Unicode码(2字节)
    Dim UtfB    ''Utf-8单个字节
    Dim UtfB1, UtfB2, UtfB3 ''Utf-8码的三个字节
    Dim i, n, s
    n = 0
    ub = 0
    For i = 1 To Len(URL)
        B = Mid(URL, i, 1)
        Select Case B
            Case "+"
                s = s & " "
            Case "%"
                ub = Mid(URL, i + 1, 2)
                UtfB = CInt("&H" & ub)
                If UtfB < 128 Then
                    i = i + 2
                    s = s & ChrW(UtfB)
                Else
                    UtfB1 = (UtfB And &HF) * &H1000   ''取第1个Utf-8字节的二进制后4位
                    UtfB2 = (CInt("&H" & Mid(URL, i + 4, 2)) And &H3F) * &H40      ''取第2个Utf-8字节的二进制后6位
                    UtfB3 = CInt("&H" & Mid(URL, i + 7, 2)) And &H3F      ''取第3个Utf-8字节的二进制后6位
                    s = s & ChrW(UtfB1 Or UtfB2 Or UtfB3)
                    i = i + 8
                End If
            Case Else    ''Ascii码
                s = s & B
        End Select
    Next
    UTF8_UrlDecode = s
End Function
'//////////////////////////////////////////////////////////////////////////////
'//
'//      继承实现
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有属性
'//
'//////////////////////////////////////////////////////////////////////////////
Property Get GapSize() As Long
  GapSize = mGapSize
End Property
Property Let GapSize(ByVal newValue As Long)
  If mCount > 0 Then
    Err.Raise 1000, , "[SunSoft]已经有数据，容量的增量不得修改"
  End If
  If newValue <= 0 Then
    Err.Raise 1001, , "[SunSoft]容量的增量不得为负数"
  End If
  mGapSize = newValue
End Property


'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有方法
'//
'//////////////////////////////////////////////////////////////////////////////
Public Function ToURLParams(ParamArray Params()) As String
  Dim param As Variant
  Dim encodedURL As String
  For Each param In Params
    If encodedURL = "" Then
      encodedURL = param & "=" & UTF8_URLEncoding(Item(param))
    Else
      encodedURL = encodedURL & "&" & param & "=" & UTF8_URLEncoding(Item(param))
    End If
  Next
  ToURLParams = encodedURL
End Function

Public Sub Add(ByVal Key As String, ByVal Value As Variant)
  Dim index As Long
  
  If findIndex(Key) > 0 Then
    Err.Raise 1002, , "[SunSoft]关键字重复，无法添加"
  End If
  index = findFreeIndex(Key)
  If index < 0 Then
    Call expandArray
    index = findFreeIndex(Key)
  End If
  mArray(index).IsTaken = True
  mArray(index).Key = Key
  mArray(index).Value = Value
End Sub

Public Property Get Item(ByVal key As String) As Variant
  Dim index As Long
  index = findIndex(key)
  Item = mArray(index).Value
End Property

Public Property Let Item(ByVal key As String, Value As Variant)
  Dim index As Long
  
  index = findIndex(key)
  
  If index > 0 Then
    mArray(index).IsTaken = True
    mArray(index).key = key
    mArray(index).Value = Value
    Exit Property
  End If
  index = findFreeIndex(key)
  If index < 0 Then
    Call expandArray
    index = findFreeIndex(key)
  End If
  mArray(index).IsTaken = True
  mArray(index).key = key
  mArray(index).Value = Value
End Property
