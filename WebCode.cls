VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private WithEvents req As WinHttp.WinHttpRequest
Attribute req.VB_VarHelpID = -1
Private canGet As Boolean
Private mTimeOut As Long
Private Const MIN_TIME_OUT = 10
Private Const MAX_TIME_OUT = 120
Private Const TIME_OUT_STEP = 10
'调整大小
Private Sub mShorterTimeOut()
  mTimeOut = mTimeOut - TIME_OUT_STEP
  If mTimeOut <= MIN_TIME_OUT Then
    mTimeOut = MIN_TIME_OUT
  End If
End Sub

Private Sub mLongerTimeOut()
  mTimeOut = mTimeOut + TIME_OUT_STEP
  If mTimeOut >= MAX_TIME_OUT Then
    mTimeOut = MIN_TIME_OUT
  End If
End Sub
'说明：此代码类模块用于获取网页的源代码
Private Function getHtmlStrGB(ByVal strURL As String) As String   '获取源码
  DoEvents
  Dim stime As Date, ntime As Date
  Dim XmlHttp As Variant
  '"获取网页源码"
  Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
  XmlHttp.Open "GET", strURL, True
  XmlHttp.SetRequestHeader "If-Modified-Since", "0"
  XmlHttp.Send
  stime = Now '获取当前时间
  While XmlHttp.readyState <> 4
    DoEvents
    ntime = Now '获取循环时间
    If DateDiff("s", stime, ntime) > mTimeOut Then
      Call mLongerTimeOut
      Err.Raise 10001, , "[SunSoft]连接超时"
    End If
  Wend
  
  If Not mIsStatusOK(XmlHttp.status) Then
    Err.Raise 10002, , "[SunSoft]连接错误"
  End If
  
  Call mShorterTimeOut

  getHtmlStrGB = StrConv(XmlHttp.ResponseBody, vbUnicode)
  Set XmlHttp = Nothing
  DoEvents
End Function

Private Function Utf8ToUnicode(ByRef Utf() As Byte) As String
  Dim utfLen As Long
  utfLen = -1
  On Error Resume Next
  utfLen = UBound(Utf)
  If utfLen = -1 Then Exit Function
  
  On Error GoTo 0
  
  Dim i As Long, j As Long, k As Long, N As Long
  Dim B As Byte, cnt As Byte
  Dim Buf() As String
  ReDim Buf(utfLen)
  
  i = 0
  j = 0
  Do While i <= utfLen
    B = Utf(i)
    
    If (B And &HFC) = &HFC Then
        cnt = 6
    ElseIf (B And &HF8) = &HF8 Then
        cnt = 5
    ElseIf (B And &HF0) = &HF0 Then
        cnt = 4
    ElseIf (B And &HE0) = &HE0 Then
        cnt = 3
    ElseIf (B And &HC0) = &HC0 Then
        cnt = 2
    Else
        cnt = 1
    End If
        
    If i + cnt - 1 > utfLen Then
        Buf(j) = "?"
        Exit Do
    End If
        
    Select Case cnt
    Case 2
        N = B And &H1F
    Case 3
        N = B And &HF
    Case 4
        N = B And &H7
    Case 5
        N = B And &H3
    Case 6
        N = B And &H1
    Case Else
        Buf(j) = Chr(B)
        GoTo Continued:
    End Select
                
        For k = 1 To cnt - 1
            B = Utf(i + k)
            N = N * &H40 + (B And &H3F)
        Next
        If N < -32768 Or N > 65535 Then
          Buf(j) = "?"
        Else
          Buf(j) = ChrW(N)
        End If
        
Continued:
        i = i + cnt
        j = j + 1
  Loop
    
  Utf8ToUnicode = Join(Buf, "")
End Function

Public Function Post(ByVal url As String, ByVal data As String) As String
  On Error Resume Next
  DoEvents
  Dim stime As Date, ntime As Date
  Dim XmlHttp As Variant
  '"获取网页源码"
  Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
  XmlHttp.Open "POST", url, True
  XmlHttp.SetRequestHeader "If-Modified-Since", "0"
  XmlHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  XmlHttp.Send data
  stime = Now '获取当前时间
  While XmlHttp.readyState <> 4
    DoEvents
    ntime = Now '获取循环时间
    If DateDiff("s", stime, ntime) > mTimeOut Then
      Call mLongerTimeOut
      Err.Raise 10001, , "[SunSoft]连接超时"
    End If
  Wend
  
  If Not mIsStatusOK(XmlHttp.status) Then
    Err.Raise 10002, , "[SunSoft]连接错误"
  End If
  
  Call mShorterTimeOut
  Post = StrConv(XmlHttp.ResponseBody, vbUnicode)
  Set XmlHttp = Nothing
  DoEvents
End Function

Private Function getHtmlStrUTF(ByVal strURL As String) As String   '获取源码
  Dim smt() As Byte

  DoEvents
  Dim stime, ntime
  Dim XmlHttp
  ' St "获取网页源码"
  Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
  XmlHttp.Open "GET", strURL, True
  XmlHttp.SetRequestHeader "If-Modified-Since", "0"
  XmlHttp.Send
  stime = Now '获取当前时间
  While XmlHttp.readyState <> 4
    DoEvents
    ntime = Now '获取循环时间
    If DateDiff("s", stime, ntime) > mTimeOut Then
      Call mLongerTimeOut
      Err.Raise 10001, , "[SunSoft]连接超时"
    End If
    DoEvents
  Wend
  
  If Not mIsStatusOK(XmlHttp.status) Then
    Err.Raise 10002, , "[SunSoft]连接错误"
  End If
  
  Call mShorterTimeOut
  If LenB(XmlHttp.ResponseBody) > 0 Then
    smt = XmlHttp.ResponseBody
  Else
    getHtmlStrUTF = ""
    Exit Function
  End If
  getHtmlStrUTF = Utf8ToUnicode(smt)
  Set XmlHttp = Nothing
  DoEvents
  Exit Function
End Function

Private Function mIsStatusOK(ByVal status As Long) As Boolean
  Select Case status
    Case 200, 201, 202
      mIsStatusOK = True
    Case Else
      mIsStatusOK = False
  End Select
End Function

Public Function GetHTMLCode(ByVal url As String, Optional ByVal charset As String = "gb2312") As String
  Dim code As String
  Select Case charset
    Case "gb2312"
      code = getHtmlStrGB(url)
    Case "utf-8"
      code = getHtmlStrUTF(url)
    Case "smart"
      code = getHtmlStrGB(url)
      code = Trim(code)
      code = LCase(code)
      If Abs(InStr(1, code, "charset") - InStr(1, code, "utf-8")) < 10 Then
        code = getHtmlStrUTF(url)
      End If
  End Select
  GetHTMLCode = code
End Function

Public Function GetHttpCode(ByVal url As String, Optional ByVal charset As String = "gb2312") As String
  Dim code As String
  Select Case charset
    Case "gb2312"
      code = GetHttpCodeGB(url)
    Case "utf-8"
      code = GetHttpCodeUTF(url)
    Case "smart"
      code = GetHttpCodeGB(url)
      code = Trim(code)
      code = LCase(code)
      If Abs(InStr(1, code, "charset") - InStr(1, code, "utf-8")) < 10 Then
          code = GetHttpCodeUTF(url)
      End If
  End Select
  GetHttpCode = code
End Function

Public Function GetHttpCodeUTF(ByVal url As String) As String
  Dim smt() As Byte
  Dim stime As Date, ntime As Date
redo:
  Set req = New WinHttp.WinHttpRequest
  
  req.Open "GET", url, True
  req.SetRequestHeader "If-Modified-Since", "0"
  req.SetRequestHeader "User-Agent", "Mozilla/6.0"
  req.Send
  If req Is Nothing Then
      GoTo redo
  End If
  stime = Now '获取当前时间
  canGet = False
  While canGet = False
      DoEvents
      ntime = Now '获取循环时间
      If DateDiff("s", stime, ntime) > mTimeOut Then
        Call mLongerTimeOut
        Err.Raise 10001, , "[SunSoft]连接超时"
      End If
  Wend
  If req Is Nothing Then
      GoTo redo
  End If

  If Not mIsStatusOK(req.status) Then
    Err.Raise 10002, , "[SunSoft]连接错误"
  End If

  Call mShorterTimeOut
  If LenB(req.ResponseBody) > 0 Then
      smt = req.ResponseBody
  Else
      GetHttpCodeUTF = ""
      Exit Function
  End If
  
  GetHttpCodeUTF = Utf8ToUnicode(smt)
  
  Set req = Nothing
End Function

Public Function GetHttpCodeGB(ByVal url As String) As String
  Dim stime As Date, ntime As Date
redo:
  Set req = New WinHttp.WinHttpRequest

  req.Open "GET", url, True
  req.SetRequestHeader "If-Modified-Since", "0"
  req.SetRequestHeader "User-Agent", "Mozilla/6.0"
  req.Send
  If req Is Nothing Then
      GoTo redo
  End If
  stime = Now '获取当前时间
  canGet = False
  While canGet = False
      DoEvents
      ntime = Now '获取循环时间
      If DateDiff("s", stime, ntime) > mTimeOut Then
        Call mLongerTimeOut
        Err.Raise 10001, , "[SunSoft]连接超时"
      End If
  Wend
  If req Is Nothing Then
      GoTo redo
  End If
    
  If Not mIsStatusOK(req.status) Then
    Err.Raise 10002, , "[SunSoft]连接错误"
  End If
    
  Call mShorterTimeOut
  If LenB(req.ResponseBody) > 0 Then
      GetHttpCodeGB = StrConv(req.ResponseBody, vbUnicode)
  Else
      GetHttpCodeGB = ""
  End If

  Set req = Nothing
End Function

Private Sub Class_Initialize()
  canGet = False
  mTimeOut = MIN_TIME_OUT
End Sub

Private Sub req_OnResponseFinished()
  canGet = True
End Sub

Public Function URLEncode(ByVal strParameter As String) As String
  Dim s As String
  Dim i As Integer
  Dim intValue As Integer
    
  Dim TmpData() As Byte
  
  s = ""
  TmpData = StrConv(strParameter, vbFromUnicode)
  For i = 0 To UBound(TmpData)
    intValue = TmpData(i)
    If (intValue >= 48 And intValue <= 57) Or _
      (intValue >= 65 And intValue <= 90) Or _
      (intValue >= 97 And intValue <= 122) Then
      s = s & Chr(intValue)
    ElseIf intValue = 32 Then
      s = s & "+"
    Else
      s = s & "%" & Hex(intValue)
    End If
  Next i
  URLEncode = s
  
End Function

Public Function Utf8UrlEncode(ByVal text As String) As String
  Dim utf8Code As String
  utf8Code = UnicodeToUtf8(text)
  Utf8UrlEncode = URLEncode(utf8Code)
End Function

Public Function UnicodeToUtf8(ByVal UCS As String) As Byte()
  Dim lLength As Long
  Dim lBufferSize As Long
  Dim lResult As Long
  Dim abUTF8() As Byte
  
  lLength = Len(UCS)
  If lLength = 0 Then Exit Function
  lBufferSize = lLength * 3 + 1
  ReDim abUTF8(lBufferSize - 1)
  lResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(UCS), lLength, abUTF8(0), lBufferSize, vbNullString, 0)
  If lResult <> 0 Then
    lResult = lResult - 1
    ReDim Preserve abUTF8(lResult)
    UnicodeToUtf8 = abUTF8
  End If
End Function

Public Function UTF8_URLEncoding(ByVal szInput As String)
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

Public Property Get TimeOut() As Variant
  TimeOut = mTimeOut
End Property

