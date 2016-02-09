VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConnProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'private
Private m_FilePath As String

'public
Public Db As New AdodbHelper

Private Sub Class_Initialize()
  m_FilePath = Replace(App.Path & "\documents.mdb", "\\", "\")
End Sub

Public Property Get DbHelper() As AdodbHelper
  Db.SetConnToFile m_FilePath
  Set DbHelper = Db
End Property

Public Property Get DbPath() As String
  DbPath = m_FilePath
End Property
