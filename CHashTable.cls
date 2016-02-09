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

Public Function Item(ByVal Key As String) As Variant
  Dim index As Long
  index = findIndex(Key)
  Item = mArray(index).Value
End Function
