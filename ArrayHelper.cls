VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayHelper"
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


'------------------------------------------------------------------------------
'       公有常量
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
'//      公有属性
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有方法
'//
'//////////////////////////////////////////////////////////////////////////////
Public Function StringArray(ParamArray arr()) As String()
  Dim a() As String
  Dim i As Integer
  ReDim a(UBound(arr))
  For i = LBound(arr) To UBound(arr)
    If VarType(arr(i)) <> vbString Then
      Err.Raise 104, , "[SunSoft]参数非String，请检查代码！"
    End If
    a(i) = CStr(arr(i))
  Next i
  StringArray = a
End Function

Public Function IntegerArray(ParamArray arr()) As Integer()
  Dim a() As Integer
  Dim i As Integer
  ReDim a(UBound(arr))
  For i = LBound(arr) To UBound(arr)
    If VarType(arr(i)) <> vbInteger Then
      Err.Raise 104, , "[SunSoft]参数非Integer，请检查代码！"
    End If
    a(i) = CStr(arr(i))
  Next i
  IntegerArray = a
End Function
