VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextBoxHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'【成员变量】
Private WithEvents m_TextBox As TextBox
Attribute m_TextBox.VB_VarHelpID = -1

Private m_NumberOnly         As Boolean

Private m_CancelKey          As Boolean

'枚举变量
Public Enum vbKeyType

    vbF5 = 116

End Enum

'【公开属性】
Public Property Get NumberOnly() As Boolean
    NumberOnly = m_NumberOnly
End Property

Public Property Let NumberOnly(ByVal vNewValue As Boolean)
    m_NumberOnly = vNewValue
End Property

Public Property Get Length() As Variant
    Length = Len(m_TextBox.Text)
End Property

'【公开方法】
'读取文本
Public Function GetText() As String
    GetText = m_TextBox.Text
End Function

'绑定文本框
Public Sub InitTextBox(ByRef txtCtl As TextBox)
    Set m_TextBox = txtCtl
End Sub

'选定所有内容
Public Sub SelAll()

    With m_TextBox
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With

End Sub

'光标移动到末尾
Public Sub PointerToEnd()
    LocPointer Length
End Sub

'清空文本框
Public Sub Clear()
    m_TextBox.Text = ""
End Sub

'获取选定的文本
Public Function GetSelected()
    GetSelected = m_TextBox.SelText
End Function

'【私有方法】
'定位光标
Private Sub LocPointer(ByVal offsetValue As Long)
    m_TextBox.SelStart = offsetValue
    m_TextBox.SetFocus
End Sub

'【事件绑定】
Private Sub m_TextBox_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim txtContent As String

    Select Case True

        Case NumberOnly
            txtContent = m_TextBox.Text

            If IsNumeric(txtContent) = False And NumberOnly Then
                m_TextBox.Text = Mid(txtContent, 1, Len(txtContent) - 1)
                PointerToEnd
            End If

        Case KeyCode = vbKeyA And Shift = vbCtrlMask
            Call SelAll
            m_CancelKey = True
    End Select

End Sub

Private Sub m_TextBox_KeyPress(KeyAscii As Integer)

    If m_CancelKey = True Then
        KeyAscii = 0
        m_CancelKey = False
    End If

End Sub

