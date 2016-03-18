VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CString"
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


'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有声明
'//
'//////////////////////////////////////////////////////////////////////////////

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


'------------------------------------------------------------------------------
'       私有数据类型
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       私有变量
'------------------------------------------------------------------------------
Private mString As String

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
  '
End Sub


'------------------------------------------------------------------------------
'       销毁
'------------------------------------------------------------------------------
Private Sub Class_Terminate()
  '
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
Private Function isArrEmpty(arr() As String) As Boolean
  On Error Resume Next
  Dim size As Long
  size = UBound(arr)
  If Err.Number > 0 Then
    isArrEmpty = True
  Else
    isArrEmpty = False
  End If
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
Public Property Get Length() As Long
  Length = Len(mString)
End Property

Public Property Let text(ByVal vNewValue As String)
Attribute text.VB_UserMemId = 0
Attribute text.VB_MemberFlags = "40"
  mString = vNewValue
End Property
'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有方法
'//
'//////////////////////////////////////////////////////////////////////////////
Public Function UTF8_URLEncoding(ByVal url As String)
  Dim wch As String, uch As String, szRet As String
  Dim x As Long
  Dim nAsc As Long, nAsc2 As Long, nAsc3 As Long
  If url = "" Then
    UTF8_URLEncoding = url
    Exit Function
  End If
  For x = 1 To Len(url)
    wch = Mid(url, x, 1)
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

Public Function UTF8_UrlDecode(ByVal url As String)
  Dim SingleWord As String, UtfBStr As String   ''中文字的Unicode码(2字节)
  Dim UtfB As Byte    ''Utf-8单个字节
  Dim UtfB1 As Long, UtfB2 As Long, UtfB3 As Long ''Utf-8码的三个字节
  Dim i As Long, OriginUrl As String

  For i = 1 To Len(url)
    SingleWord = Mid(url, i, 1)
    Select Case SingleWord
      Case "+"
        OriginUrl = OriginUrl & " "
      Case "%"
        UtfBStr = Mid(url, i + 1, 2)
        UtfB = CInt("&H" & UtfBStr)
        If UtfB < 128 Then
          i = i + 2
          OriginUrl = OriginUrl & ChrW(UtfB)
        Else
          UtfB1 = CLng(UtfB And &HF) * &H1000   ''取第1个Utf-8字节的二进制后4位
          UtfB2 = (CInt("&H" & Mid(url, i + 4, 2)) And &H3F) * &H40      ''取第2个Utf-8字节的二进制后6位
          UtfB3 = CInt("&H" & Mid(url, i + 7, 2)) And &H3F      ''取第3个Utf-8字节的二进制后6位
          OriginUrl = OriginUrl & ChrW(UtfB1 Or UtfB2 Or UtfB3)
          i = i + 8
        End If
      Case Else    ''Ascii码
        OriginUrl = OriginUrl & SingleWord
    End Select
  Next
  UTF8_UrlDecode = OriginUrl
End Function

Public Function Equals(ByRef cs As CString) As Boolean
  Equals = (cs.ToString = mString)
End Function

Public Function ToString() As String
  ToString = mString
End Function
'返回的数据是1基的
Public Function IndexOf(ByVal value As String) As Long
  IndexOf = InStr(1, mString, value)
End Function

Public Function LastIndexOf(ByVal value As String) As Long
  LastIndexOf = InStrRev(mString, value)
End Function

Public Function ToByteArray() As Byte()
  
End Function

Public Function IsNullOrEmpty(ByRef text As String) As Boolean
  If text = "" Or text = vbNullString Then
    IsNullOrEmpty = True
  Else
    IsNullOrEmpty = False
  End If
End Function

Public Function Join(ByVal Separator As String, StrArr() As String) As String
  Dim joinedText As String, i As Long
  If isArrEmpty(StrArr) Then
    Join = ""
    Exit Function
  End If
  For i = LBound(StrArr) To UBound(StrArr)
    If i = LBound(StrArr) Then
      joinedText = StrArr(i)
    Else
      joinedText = joinedText & Separator & StrArr(i)
    End If
  Next i
  Join = joinedText
End Function


