VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VB_Cookie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'require:
'VB_ExCollection
Option Explicit
'Private Variables
Private cookies As VB_ExCollection
'Init Class
Private Sub Class_Initialize()
  Set cookies = New VB_ExCollection
End Sub

Private Sub Class_Terminate()
  Set cookies = Nothing
End Sub
'Private Functions
'private functions
Private Function getCookieData(ByVal cookieStr As String) As VB_Cookie_Data
  'p_skey=mYLws2hzVCzwmkJP2hlDNpQjsFyQWSVrPta0H70wvcQ_; PATH=/; DOMAIN=web2.qq.com;
  Dim cookieArr() As String
  Dim pCD As New VB_Cookie_Data
  
  cookieArr = Split(cookieStr, ";")
  pCD.cookieString = Trim(cookieArr(0))
  
  pCD.key = Split(pCD.cookieString, "=")(0)
  pCD.value = Split(pCD.cookieString, "=")(1)
  
  If InStr(1, cookieStr, "DOMAIN") > 0 Then
    pCD.domain = Split(Split(cookieStr, "DOMAIN=")(1), ";")(0)
  Else
    pCD.domain = ""
  End If
  
  If InStr(1, cookieStr, "PATH") > 0 Then
    pCD.path = Split(Split(cookieStr, "PATH=")(1), ";")(0)
  Else
    pCD.path = ""
  End If
  
  Set getCookieData = pCD
End Function
'Public Functions
Public Sub AddCookiesFromHeaders(ByVal respHeaders As String)
  Dim tmpArr() As String, cookieStr As String, cookieStrs() As String
  Dim I As Integer, cookieContentLoc As Integer, j As Integer
  Dim cookieData As VB_Cookie_Data
  tmpArr = Split(respHeaders, vbCrLf)
  For I = 0 To UBound(tmpArr)
    cookieStr = tmpArr(I)
    cookieContentLoc = InStr(1, cookieStr, "Set-Cookie:")
    If cookieContentLoc > 0 Then
      cookieContentLoc = cookieContentLoc + Len("Set-Cookie:")
      Set cookieData = getCookieData(Mid(cookieStr, cookieContentLoc))
      cookies.SetItem cookieData.key, cookieData
    End If
  Next I
End Sub

Public Function getCookie(ByVal cookieName As String) As Variant
  Set getCookie = cookies.SafeGetItem(cookieName)
End Function

Public Function cookieString() As String
  cookieString = getCookieOfDomain("")
End Function

Public Function getCookieOfDomain(Optional ByVal domain As String = "") As String
  Dim keys() As String
  Dim cookiesOfDomain As String
  Dim cookieItem As VB_Cookie_Data
  keys = cookies.AllKeys
  
  For Each key In keys
    Set cookieItem = cookies.SafeGetItem(key)
    If cookieItem.domain = domain Or cookieItem.domain = "" Then
      If cookiesOfDomain = "" Then
        cookiesOfDomain = cookieItem.cookieString
      Else
        cookiesOfDomain = cookiesOfDomain & ";" & cookieItem.cookieString
      End If
    End If
  Next key
  getCookieOfDomain = cookiesOfDomain
End Function

