VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5925
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton btnCopyMemory 
      Caption         =   "CopyMem"
      Height          =   435
      Left            =   4680
      TabIndex        =   9
      Top             =   1260
      Width           =   1155
   End
   Begin VB.CommandButton btnCArray 
      Caption         =   "CHashTable"
      Height          =   375
      Left            =   3900
      TabIndex        =   8
      Top             =   600
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "77"
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton btnNewTable 
      Caption         =   "NewTable"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton btnArray 
      Caption         =   "数组赋值"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdDelFile 
      Caption         =   "DelFile"
      Height          =   360
      Left            =   1680
      TabIndex        =   4
      Top             =   600
      Width           =   990
   End
   Begin VB.CommandButton cmdReadFile 
      Caption         =   "ReadFile"
      Height          =   360
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   990
   End
   Begin VB.TextBox txtOut 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1200
      Width           =   4455
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "新增一行"
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   990
   End
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "追加一句话"
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
     ByVal Destination As Long, _
     ByVal Source As Long, _
     ByVal Length As Long)
Private mF As New cFile

Private Sub btnArray_Click()
  Dim aString() As String
  Dim bInteger() As Integer
  Dim ah As New ArrayHelper
  aString = ah.StringArray("we are the world", "man in the mirror", "dangerous")
  bInteger = ah.IntegerArray(1, 2, 3, 4, 5)

End Sub

Private Sub btnCArray_Click()
  Dim arr As New CArray
  arr.Add "sunrui", 26
  arr.Add "sunrui8", "love"
  MsgBox arr.Item("sunrui") & arr.Item("sunrui8")
  
End Sub

Private Sub btnCopyMemory_Click()
  Dim h As New ArrayHelper
  Dim a() As String
  Dim b() As String
  Dim c() As String
  a = h.StringArray("we are the world", "beat it")
  b = h.StringArray("heal the world", "man in the mirror")
  Stop
End Sub

Private Sub btnNewTable_Click()
  Dim dbc As New DbCreateHelper
  dbc.SetDbFile "lover.mdb"
  dbc.CreateTable "Users", "id:integer,name:string,age:integer,school:string,profile:text,created_at:date"
End Sub

Private Sub cmdAddNew_Click()
    mF.WriteLineToTextFile "1.txt", "2f"
End Sub

Private Sub cmdCommand1_Click()
    mF.WriteToTextFile "1.txt", "ff"
End Sub

Private Sub cmdDelFile_Click()
    mF.Delete "1.txt"
End Sub

Private Sub cmdReadFile_Click()
    txtOut.Text = mF.ReadTextFile("1.txt")
End Sub

Private Sub Command1_Click()
  Dim k As ClassBase
  Set k = New ClassBase
  MsgBox k.Counts("123")
End Sub

Private Sub Form_Load()
  Call btnNewTable_Click
End Sub
