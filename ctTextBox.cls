VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ctTextBox"
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
Public Enum vbKeyType
  vbF1 = 112
  vbF2 = 113
  vbF3 = 114
  vbF4 = 115
  vbF5 = 116
  vbF6 = 117
  vbF7 = 118
  vbF8 = 119
  vbF9 = 120
  vbF10 = 121
  vbF11 = 122
  vbF12 = 123
End Enum

'------------------------------------------------------------------------------
'       公有变量
'------------------------------------------------------------------------------
Public CompareMethod As VbCompareMethod

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
Private WithEvents m_TextBox As TextBox
Attribute m_TextBox.VB_VarHelpID = -1
'------------------------------------------------------------------------------
'       属性变量
'------------------------------------------------------------------------------
Private m_NumberOnly As Boolean
Private m_CancelKey As Boolean
Private m_AllowCtrlA As Boolean
Private m_AllowCtrlF As Boolean
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
  '初始化：文本框支持全选
  m_AllowCtrlA = True
  m_AllowCtrlF = False
  CompareMethod = vbBinaryCompare
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
Private Sub m_TextBox_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim txtContent As String
  If m_NumberOnly Then
    If m_TextBox.text = "" Then Exit Sub
    txtContent = m_TextBox.text
    If IsNumeric(txtContent) = False And m_NumberOnly Then
      m_TextBox.text = Mid(txtContent, 1, Len(txtContent) - 1)
      PointerToEnd
    End If
  End If
End Sub

Private Sub m_TextBox_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim mSearchText As String, mSearchIndex As Long
  
  If KeyCode = vbF3 And Shift = 0 Then
    Call m_SelNextMatch
  End If
  
  If m_AllowCtrlA Then
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
      Call SelAll
      m_CancelKey = True
    End If
  End If
  
  If m_AllowCtrlF Then
    If KeyCode = vbKeyF And Shift = vbCtrlMask Then
      mSearchText = InputBox("请输入要查找的内容", "文本搜索", m_TextBox.SelText)
      If mSearchText <> "" Then
        If mSearchText = m_TextBox.SelText Then
          'mSearchIndex = InStr(m_TextBox.SelStart + 2, m_TextBox.text, mSearchText, CompareMethod)
          Call m_SelNextMatch
        Else
          mSearchIndex = InStr(m_TextBox.SelStart + 1, m_TextBox.text, mSearchText, CompareMethod)
        End If
        If mSearchIndex > 0 Then
          mSearchIndex = mSearchIndex - 1
          With m_TextBox
            .SelStart = mSearchIndex
            .SelLength = Len(mSearchText)
            .SetFocus
          End With
        End If
      End If
      m_CancelKey = True
    End If
  End If
End Sub

Private Sub m_TextBox_KeyPress(KeyAscii As Integer)
  Dim txtContent As String
  Debug.Print "KeyPress"
  If m_CancelKey = True Then
    KeyAscii = 0
    m_CancelKey = False
  End If

  If m_NumberOnly Then
    If m_TextBox.text = "" Then Exit Sub
    txtContent = m_TextBox.text
    If IsNumeric(txtContent) = False And m_NumberOnly Then
      m_TextBox.text = Mid(txtContent, 1, Len(txtContent) - 1)
      PointerToEnd
    End If
  End If
End Sub

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
Private Sub LocPointer(ByVal offsetValue As Long)
  m_TextBox.SelStart = offsetValue
  m_TextBox.SetFocus
End Sub

Private Sub m_SelNextMatch()
  Dim mSelText As String
  Dim mSelIndex As Long
  mSelText = m_TextBox.SelText
  mSelIndex = m_TextBox.SelStart
  If mSelText = "" Then Exit Sub
  mSelIndex = InStr(mSelIndex + 2, m_TextBox.text, mSelText, CompareMethod)
  If mSelIndex = 0 Then
    'cannot find from here
    mSelIndex = InStr(1, m_TextBox.text, mSelText, CompareMethod)
    If mSelIndex = 0 Then
      'never can find
    Else
      With m_TextBox
        .SelStart = mSelIndex - 1
        .SelLength = Len(mSelText)
        .SetFocus
      End With
    End If
  Else
    'found
    With m_TextBox
      .SelStart = mSelIndex - 1
      .SelLength = Len(mSelText)
      .SetFocus
    End With
  End If
End Sub

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
Public Property Get NumberOnly() As Boolean
  NumberOnly = m_NumberOnly
End Property
Public Property Let NumberOnly(ByVal vNewValue As Boolean)
  m_NumberOnly = vNewValue
End Property

Public Property Get Length() As Variant
  Length = Len(m_TextBox.text)
End Property

Public Property Get AllowCtrlA() As Boolean
  AllowCtrlA = m_AllowCtrlA
End Property
Public Property Let AllowCtrlA(ByVal vNewValue As Boolean)
  m_AllowCtrlA = vNewValue
End Property

Public Property Get AllowCtrlF() As Boolean
  AllowCtrlF = m_AllowCtrlF
End Property
Public Property Let AllowCtrlF(ByVal vNewValue As Boolean)
  m_AllowCtrlF = vNewValue
End Property

'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有方法
'//
'//////////////////////////////////////////////////////////////////////////////
Public Function GetText() As String
  GetText = m_TextBox.text
End Function

Public Sub Bind(ByRef txtCtl As TextBox)
  Set m_TextBox = txtCtl
End Sub

Public Sub SelAll()
  With m_TextBox
    .SelStart = 0
    .SelLength = Len(.text)
    .SetFocus
  End With
End Sub

Public Sub PointerToEnd()
  LocPointer Length
End Sub

Public Function GetSelected()
  GetSelected = m_TextBox.SelText
End Function


