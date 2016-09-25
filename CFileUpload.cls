VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileUpload"
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
'//      ��������
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       �ӿڼ̳�
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       ���г���
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       ������������
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       ���б���
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       ����API
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       �¼�����
'------------------------------------------------------------------------------


'//////////////////////////////////////////////////////////////////////////////
'//
'//      ˽������
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       ˽�г���
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       ˽����������
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       ˽�б���
'------------------------------------------------------------------------------
Private mWinHttpReq As WinHttp.WinHttpRequest '����һ������
Private mList As CList

'------------------------------------------------------------------------------
'       ���Ա���
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       ˽��API
'------------------------------------------------------------------------------


'//////////////////////////////////////////////////////////////////////////////
'//
'//      ��
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       ��ʼ��
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
  Set mWinHttpReq = New WinHttp.WinHttpRequest
  Set mList = New CList
End Sub


'------------------------------------------------------------------------------
'       ����
'------------------------------------------------------------------------------
Private Sub Class_Terminate()
  Set mWinHttpReq = Nothing
  Set mList = Nothing
End Sub


'//////////////////////////////////////////////////////////////////////////////
'//
'//      �¼�����
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      ˽������
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      ˽�з���
'//
'//////////////////////////////////////////////////////////////////////////////
Private Function LoadData(ByVal FilePath As String) As Byte()
'�����ļ����ֽ�����
  Dim Bin() As Byte, FileLength As Long
  Open FilePath For Binary As #1                   '�Զ����Ʒ�ʽ���ļ�
    FileLength = LOF(1) ''��ȡ�ļ����ȡ��ô�С���ֽ�Ϊ��λ
    ReDim Bin(FileLength - 1)                       '����Ϊ��̬����������·���洢�ռ䡣
    Get #1, , Bin                                 '��һ���Ѵ򿪵Ĵ����ļ�����һ������֮�С�
  Close #1
  LoadData = Bin
End Function

Private Function mGetFileName(ByVal mFilePath As String) As String
'���ļ�·���л�ȡ�ļ���
  mFilePath = Replace(mFilePath, "\\", "\")
  mGetFileName = Mid(mFilePath, InStrRev(mFilePath, "\") + 1)
End Function

Private Function BinString(ByVal content As String) As Byte()
'��ȡ�ַ������ֽ�����
  Dim source() As Byte, i As Long, countUp As Long, dest() As Byte
  source = content
  countUp = 0
  
  If Len(content) = 0 Then
    Exit Function
  End If
  
  ReDim dest(UBound(source))
  
  For i = 0 To UBound(source)
    If source(i) <> 0 Then
      dest(countUp) = source(i)
      countUp = countUp + 1
    End If
  Next i
  
  ReDim Preserve dest(countUp - 1)
  BinString = dest
End Function

Private Function mGetHost(ByVal mUrl As String) As String
'http://127.0.0.2:80/
  Dim beginPos As Long, endPos As Long
  beginPos = InStr(1, mUrl, "//")
  If beginPos + 2 > Len(mUrl) Then
    mGetHost = ""
    Exit Function
  Else
    beginPos = beginPos + 2
  End If
  endPos = InStr(beginPos, mUrl, "/")
  If endPos = 0 Then endPos = Len(mUrl) + 1
  mGetHost = Mid(mUrl, beginPos, endPos - beginPos)
End Function

Private Sub mJoinBin(dest() As Byte, source() As Byte)
'�ϳ��ֽ�����
  Dim baseSize As Long, i As Long
  
  If UBound(dest) > 0 Then
    baseSize = UBound(dest) + 1
    ReDim Preserve dest(UBound(dest) + UBound(source) + 1)
  Else
    baseSize = 0
    ReDim dest(UBound(source))
  End If
  For i = 0 To UBound(source)
    dest(baseSize + i) = source(i)
  Next i
End Sub

Private Function BytesToBstr(strBody As Variant, ByVal CodeBase As String)
'����ת��("UTF-8"����"GB2312"����"GBK")
  Dim ObjStream
  Set ObjStream = CreateObject("Adodb.Stream")
  With ObjStream
    .Type = 1
    .Mode = 3
    .Open
    .Write strBody
    .position = 0
    .Type = 2
    .charset = CodeBase
    BytesToBstr = .ReadText
    .Close
  End With
  Set ObjStream = Nothing
End Function
'//////////////////////////////////////////////////////////////////////////////
'//
'//      �̳�ʵ��
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      ��������
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      ���з���
'//
'//////////////////////////////////////////////////////////////////////////////
Public Function UploadFile(ByVal URL As String, ByVal FilePath As String, Optional ByVal filename As String = "") As String
  On Error GoTo lastsave
'��ȡ��֤��ͼƬ������
  Dim boundary As String
  Dim part1 As String, part3 As String
  Dim data() As Byte, bindata() As Byte
  
  boundary = "------------------------------831b7f6f6a2d"
  ReDim bindata(0)
  
  data = LoadData(FilePath)
  If filename = "" Then
    filename = mGetFileName(FilePath)
  End If

  With mList
    .Clear
    .Add boundary & vbCrLf
    .Add "Content-Disposition: form-data; name=""upfile""; filename=""" & filename & """" & vbCrLf
    .Add "Content-Type: application/octet-stream" & vbCrLf
    .Add "" & vbCrLf
    part1 = .Join("")

    .Clear
    .Add "" & vbCrLf
    .Add boundary & "--" & vbCrLf
    part3 = .Join("")
  End With
  
  mJoinBin bindata, BinString(part1)
  mJoinBin bindata, data
  mJoinBin bindata, BinString(part3)
  mWinHttpReq.Open "POST", URL, True
  mWinHttpReq.SetTimeouts 30000, 30000, 30000, 30000
  mWinHttpReq.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 5.1; zh-CN; rv:1.9.1.3) Gecko/20100101 Firefox/8.0"
  mWinHttpReq.SetRequestHeader "Host", mGetHost(URL)
  mWinHttpReq.SetRequestHeader "Accept", "*/*"
  mWinHttpReq.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & Mid(boundary, 3)

  mWinHttpReq.Send bindata        '����
  mWinHttpReq.WaitForResponse  '�첽����
  UploadFile = BytesToBstr(mWinHttpReq.ResponseBody, "UTF-8")
  Exit Function
lastsave:
  Err.Clear
  Debug.Print Err.description
  UploadFile = "network-error"
End Function