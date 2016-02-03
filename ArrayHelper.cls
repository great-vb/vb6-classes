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
'       ���г���
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
    '
End Sub


'------------------------------------------------------------------------------
'       ����
'------------------------------------------------------------------------------
Private Sub Class_Terminate()
    '
End Sub

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
Public Function StringArray(ParamArray arr()) As String()
  Dim a() As String
  Dim i As Integer
  ReDim a(UBound(arr))
  For i = LBound(arr) To UBound(arr)
    If VarType(arr(i)) <> vbString Then
      Err.Raise 104, , "[SunSoft]������String��������룡"
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
      Err.Raise 104, , "[SunSoft]������Integer��������룡"
    End If
    a(i) = CStr(arr(i))
  Next i
  IntegerArray = a
End Function