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
Private WithEvents req As WinHttp.WinHttpRequest
Attribute req.VB_VarHelpID = -1
Private canGet As Boolean
'说明：此代码类模块用于获取网页的源代码
Private Function getHtmlStrGB(ByVal strURL As String) As String   '获取源码
    On Error GoTo reStart
reStart:
    DoEvents
    Dim stime As Date, ntime As Date
    Dim XmlHttp As Variant
    '"获取网页源码"
    Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
    XmlHttp.Open "GET", strURL, True
    XmlHttp.SetRequestHeader "If-Modified-Since", "0"
    XmlHttp.Send
    stime = Now '获取当前时间
    While XmlHttp.ReadyState <> 4
        DoEvents
        ntime = Now '获取循环时间
        If DateDiff("s", stime, ntime) > 10 Then getHtmlStrGB = "OutTime": Exit Function '判断超出3秒即超时退出过程
    Wend
    getHtmlStrGB = StrConv(XmlHttp.ResponseBody, vbUnicode)
    If getHtmlStrGB = "" Then getHtmlStrGB = "OutTime"
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
        
        Buf(j) = ChrW(N)
Continued:
        i = i + cnt
        j = j + 1
    Loop
    
    Utf8ToUnicode = Join(Buf, "")
 End Function
Private Function getHtmlStrUTF(ByVal strURL As String) As String   '获取源码
    On Error GoTo reStart
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
    While XmlHttp.ReadyState <> 4
        DoEvents
        ntime = Now '获取循环时间
        If DateDiff("s", stime, ntime) > 5 Then getHtmlStrUTF = "OutTime": Exit Function  '判断超出3秒即超时退出过程
        DoEvents
    Wend
    If LenB(XmlHttp.ResponseBody) > 0 Then
        smt = XmlHttp.ResponseBody
    Else
        getHtmlStrUTF = ""
        Exit Function
    End If
    getHtmlStrUTF = Utf8ToUnicode(smt)
    If getHtmlStrUTF = "" Then getHtmlStrUTF = "OutTime"
    Set XmlHttp = Nothing
    DoEvents
    Exit Function
reStart:
    getHtmlStrUTF = ""
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

Public Function GetByDiv(ByVal code As String, ByVal divBegin As String, divEnd As String) As String   '获取分隔符所夹的内容[完成，未测试]
    Dim lgStart As Long
    Dim lens As Long
    Dim lgEnd As Long
    lens = Len(divBegin)
    If InStr(1, code, divBegin) = 0 Then
        GetByDiv = ""
        Exit Function
    End If
    lgStart = InStr(1, code, divBegin) + CLng(lens)
    lgEnd = InStr(lgStart, code, divEnd)
    If lgEnd < lgStart Then
        GetByDiv = ""
        Exit Function
    End If
    GetByDiv = Mid(code, lgStart, lgEnd - lgStart)
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
        If DateDiff("s", stime, ntime) > 5 Then GetHttpCodeUTF = "": Exit Function  '判断超出3秒即超时退出过程
    Wend
    If req Is Nothing Then
        GoTo redo
    End If
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
        If DateDiff("s", stime, ntime) > 5 Then GetHttpCodeGB = "": Exit Function  '判断超出3秒即超时退出过程
    Wend
    If req Is Nothing Then
        GoTo redo
    End If
    If LenB(req.ResponseBody) > 0 Then
        GetHttpCodeGB = StrConv(req.ResponseBody, vbUnicode)
    Else
        GetHttpCodeGB = ""
    End If

    Set req = Nothing
End Function

Private Sub Class_Initialize()
    canGet = False
End Sub

Private Sub req_OnResponseFinished()
    canGet = True
End Sub


