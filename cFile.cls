VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : cFile
'    Project    : Recent
'
'    Description: 通用文件访问
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Private Declare Function PathFileExists _
                Lib "shlwapi.dll" _
                Alias "PathFileExistsA" (ByVal pszPath As String) As Boolean

Private Declare Function ShellExecute _
                Lib "shell32.dll" _
                Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                       ByVal lpOperation As String, _
                                       ByVal lpFile As String, _
                                       ByVal lpParameters As String, _
                                       ByVal lpDirectory As String, _
                                       ByVal nShowCmd As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long


'【初始化】
Private Sub Class_Initialize()

End Sub

'【公共函数：写文件】
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       Recent
' Procedure  :       WriteToTextFile
' Description:       向文件写入文本
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-14:30:32
'
' Parameters :       FilePath (String)
'                    Content (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function WriteToTextFile(ByVal FilePath As String, _
                                ByVal Content As String) As Boolean

    On Error Resume Next

    Dim fileId As Long

    fileId = FreeFile

    If PathFileExists(FilePath) Then
        Open FilePath For Append As fileId
    Else
        Open FilePath For Output As fileId
    End If

    Print #fileId, Content;
    Close fileId

    If Err.Number > 0 Then
        WriteToTextFile = False
        Err.Clear
    Else
        WriteToTextFile = True
    End If

End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       Recent
' Procedure  :       WriteLineToTextFile
' Description:       向文件写入新的一行文本
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-14:29:53
'
' Parameters :       FilePath (String)
'                    Content (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function WriteLineToTextFile(ByVal FilePath As String, _
                                    ByVal Content As String) As Boolean
    WriteLineToTextFile = WriteToTextFile(FilePath, vbCrLf & Content)
End Function

'【公共函数：读文件】

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       ReadTextFile
' Description:       读取指定路径的全文
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-15:14:52
'
' Parameters :       FilePath (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ReadTextFile(ByVal FilePath As String) As String

    Dim fileId As Long

    Dim fBin() As Byte

    Dim fSize  As Long

    fileId = FreeFile

    If FileExists(FilePath) = False Then
        Err.Raise 100, , "文件不存在！"
    End If
    
    fSize = FileLen(FilePath)
    ReDim fBin(fSize - 1)
    
    Open FilePath For Binary Access Read As fileId
        Get #fileId, , fBin
    Close fileId

    If Err.Number > 0 Then
        MsgBox Err.Description
        ReadTextFile = ""
    Else
        ReadTextFile = StrConv(fBin, vbUnicode)
    End If

End Function
 '【公共函数：文件操作】
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       FileExists
' Description:       返回指定路径是否存在文件（夹）
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-15:16:31
'
' Parameters :       FilePath (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function FileExists(ByVal FilePath As String) As Boolean
    FileExists = PathFileExists(FilePath)
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       Delete
' Description:       删除指定文件
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-16:59:10
'
' Parameters :       FilePath (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub Delete(ByVal FilePath As String)
    If FileExists(FilePath) Then
        If DeleteFile(FilePath) = 0 Then
            Err.Raise 11, , "文件删除错误！"
        End If
    Else
        Err.Raise 12, , "文件路径不存在！"
    End If
End Sub
