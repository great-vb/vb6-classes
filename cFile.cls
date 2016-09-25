VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFile"
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


'------------------------------------------------------------------------------
'       私有数据类型
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       私有变量
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       属性变量
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       私有API
'------------------------------------------------------------------------------
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
Private Declare Function DeleteFile _
                Lib "kernel32" _
                Alias "DeleteFileA" (ByVal lpFileName As String) As Long

'//////////////////////////////////////////////////////////////////////////////
'//
'//      类
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       初始化
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
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


'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有方法
'//
'//////////////////////////////////////////////////////////////////////////////
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

Public Function OverWriteToTextFile(ByVal FilePath As String, _
                                ByVal Content As String) As Boolean
    On Error Resume Next

    Dim fileId As Long

    fileId = FreeFile

    Open FilePath For Output As fileId
      Print #fileId, Content;
    Close fileId

    If Err.Number > 0 Then
        OverWriteToTextFile = False
        Err.Clear
    Else
        OverWriteToTextFile = True
    End If

End Function

Public Function WriteLineToTextFile(ByVal FilePath As String, _
                                    ByVal Content As String) As Boolean
  WriteLineToTextFile = WriteToTextFile(FilePath, vbCrLf & Content)
End Function

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

Public Function FileExists(ByVal FilePath As String) As Boolean
  FileExists = PathFileExists(FilePath)
End Function

Public Sub Delete(ByVal FilePath As String)
  If FileExists(FilePath) Then
    If DeleteFile(FilePath) = 0 Then
      Err.Raise 11, , "文件删除错误！"
    End If
  Else
    Err.Raise 12, , "文件路径不存在！"
  End If
End Sub

Public Sub Touch(ByVal FilePath As String)
  Dim fileNo As Integer
  fileNo = FreeFile
  Open FilePath For Random As fileNo
  Close fileNo
End Sub

Public Sub CreateFolder(ByVal FilePath As String)
  Dim arr() As String, i As Long, curPath As String
  FilePath = Replace(FilePath, "\\", "\")
  If Dir(FilePath, vbDirectory) = "" Then
    FilePath = Mid(FilePath, 1, InStrRev(FilePath, "\"))
  End If
  arr = Split(FilePath, "\")
  For i = 0 To UBound(arr)
    If curPath = "" Then
      curPath = curPath & arr(i)
    Else
      curPath = curPath & "\" & arr(i)
    End If
    If Dir(curPath, vbDirectory) = "" Then
      MkDir curPath
    End If
  Next i
End Sub

Public Sub SaveAsUtf8Bom(ByVal Text As String, ByVal FileName As String)
  Const adTypeText = 2
  Const adSaveCreateOverWrite = 2
  Dim oStream As Variant
  Set oStream = CreateObject("ADODB.Stream")
  oStream.Open
  oStream.charset = "UTF-8"
  oStream.Type = adTypeText
  oStream.WriteText Text
  oStream.SaveToFile FileName, adSaveCreateOverWrite
  oStream.Close
  Set oStream = Nothing
End Sub

Public Sub SaveAsUTF8(ByVal Text As String, ByVal FileName As String)
  Const adSaveCreateNotExist = 1
  Const adSaveCreateOverWrite = 2
  Const adTypeBinary = 1
  Const adTypeText = 2
  
  Dim objStreamUTF8: Set objStreamUTF8 = CreateObject("ADODB.Stream")
  Dim objStreamUTF8NoBOM: Set objStreamUTF8NoBOM = CreateObject("ADODB.Stream")
  
  With objStreamUTF8
    .charset = "UTF-8"
    .Open
    .WriteText Text
    .position = 0
    .SaveToFile FileName, adSaveCreateOverWrite
    .Type = adTypeBinary
    .position = 3
  End With
  
  With objStreamUTF8NoBOM
    .Type = adTypeBinary
    .Open
    objStreamUTF8.CopyTo objStreamUTF8NoBOM
    .SaveToFile FileName, adSaveCreateOverWrite
  End With
  
  objStreamUTF8.Close
  objStreamUTF8NoBOM.Close
End Sub


