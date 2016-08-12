VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccessHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//////////////////////////////////////////////////////////////////////////////
'@@summary
'@@require
'---Class:CHashTable.cls
'---Import:Microsoft ActiveX Data Objects 2.8 Library
'@@reference
'@@license
'@@author
'@@create
'@@modify
'---20160812:create this class
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      Public Declare
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       Interface
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       Public Const
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       Public DataType
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       Public Variable
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       Public API
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       Event Declare
'------------------------------------------------------------------------------


'//////////////////////////////////////////////////////////////////////////////
'//
'//      Private Declare
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       Private Const
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       Private DataType
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       Private Variable
'------------------------------------------------------------------------------
Private m_Conn       As ADODB.Connection

Private m_Command    As ADODB.Command

Private m_ConnString As String

Private m_FilePath   As String

Private m_Me         As New AdodbHelper

'------------------------------------------------------------------------------
'       Property Variable
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       Private API
'------------------------------------------------------------------------------


'//////////////////////////////////////////////////////////////////////////////
'//
'//      Class
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       Initialize
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
    '
End Sub


'------------------------------------------------------------------------------
'       Terminate
'------------------------------------------------------------------------------
Private Sub Class_Terminate()
  Set m_Conn = Nothing
  Set m_Command = Nothing
End Sub


'//////////////////////////////////////////////////////////////////////////////
'//
'//      Events
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      Private Property
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      Private Methods
'//
'//////////////////////////////////////////////////////////////////////////////

Private Sub OpenConn()
  Set m_Conn = New ADODB.Connection
  m_Conn.CursorLocation = adUseClient
  m_Conn.Open ConnectionString
End Sub

Private Sub CloseConn()
  m_Conn.Close
  Set m_Conn = Nothing
End Sub

Private Function m_ApostropheCount(ByVal SQL As String) As Long
'count number of "'"
  m_ApostropheCount = Len(SQL) - Len(Replace(SQL, "'", ""))
End Function

Private Function m_ProcessNameParams(mSql As String, mDic As CHashTable, mParams() As Variant) As Boolean
  Dim mNewSql As String, mWord As String, mFieldName As String
  Dim mParamCount As Long, i As Long, comaCount As Long
  Dim mBeginParam As Boolean
  
  If m_ApostropheCount(mSql) Mod 2 = 1 Then
    Err.Raise 110000000, "Symbal " '" must be in pairs,please check SQL statement"
  End If
  
  'init mDic
  mBeginParam = False
  mFieldName = ""
  mParamCount = 0
  
  For i = 1 To Len(mSql)
    mWord = Mid(mSql, i, 1)
    Select Case mWord
      Case " ", ",", ")"
        mNewSql = mNewSql & mWord
        If mBeginParam Then
          ReDim Preserve mParams(mParamCount)
          mParams(mParamCount) = mDic.Item(mFieldName)
          mFieldName = ""
          mParamCount = mParamCount + 1
        End If
        mBeginParam = False
      Case "'"
        comaCount = comaCount + 1
        mNewSql = mNewSql & mWord
      Case "@"
        If comaCount Mod 2 = 0 Then
          mBeginParam = True
          mNewSql = mNewSql & "?"
        Else
          'odd number of "'" means that "@" is only string of content
          mNewSql = mNewSql & mWord
        End If
      Case Else
        If mBeginParam = False Then
          mNewSql = mNewSql & mWord
        Else
          mFieldName = mFieldName & mWord
        End If
    End Select
  Next i
  'all done but check last word for that last word maybe param
  If mFieldName <> "" Then
    ReDim Preserve mParams(mParamCount)
    mParams(mParamCount) = mDic.Item(mFieldName)
    mFieldName = ""
  End If
  'return
  mSql = mNewSql
  m_ProcessNameParams = True
End Function

Private Function m_GetVarType(ByRef Value As Variant) As ADODB.DataTypeEnum
  Select Case VarType(Value)
    Case VbVarType.vbString
      m_GetVarType = ADODB.DataTypeEnum.adVarWChar
    Case VbVarType.vbInteger
      m_GetVarType = ADODB.DataTypeEnum.adSmallInt
    Case VbVarType.vbBoolean
      m_GetVarType = ADODB.DataTypeEnum.adBoolean
    Case VbVarType.vbCurrency
      m_GetVarType = ADODB.DataTypeEnum.adCurrency
    Case VbVarType.vbDate
      m_GetVarType = ADODB.DataTypeEnum.adDate
    Case 8209
      m_GetVarType = ADODB.DataTypeEnum.adLongVarBinary
    Case Else
      m_GetVarType = ADODB.DataTypeEnum.adVariant
  End Select
End Function

'//////////////////////////////////////////////////////////////////////////////
'//
'//      Inherit
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      Public Property
'//
'//////////////////////////////////////////////////////////////////////////////
Public Property Get ConnectionString() As String
  ConnectionString = m_ConnString
End Property

Public Property Let ConnectionString(ByVal vNewValue As String)
  m_ConnString = vNewValue
End Property

Public Property Get IsReady() As Boolean
  IsReady = IIf(Len(ConnectionString) > 0, True, False)
End Property

'//////////////////////////////////////////////////////////////////////////////
'//
'//      Public Methods
'//
'//////////////////////////////////////////////////////////////////////////////
'---------------------Data Base Connection
Public Function DbConnFromFile(ByVal FilePath As String) As ADODB.Connection
  Dim mConn As New ADODB.Connection

  mConn.CursorLocation = adUseClient
  mConn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & FilePath & ";"
  Set DbConnFromFile = mConn
End Function

Public Sub SetConnToFile(ByVal FilePath As String)
  m_ConnString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & FilePath & ";"
End Sub

Public Sub SetConnToAccdb(ByVal FilePath As String)
  m_ConnString = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=" & FilePath & ";Persist Security Info=False"
End Sub
'---------------------Querys
Public Function ExecQuery(ByVal SQL As String) As ADODB.Recordset
  Dim mRes As New ADODB.Recordset

  Set m_Command = New ADODB.Command
  Call OpenConn
  m_Command.ActiveConnection = m_Conn
  m_Command.CommandText = SQL
  Set mRes = m_Command.Execute()
  'disconnect from database
  mRes.ActiveConnection = Nothing
  Call CloseConn
  Set ExecQuery = mRes
  Set m_Command = Nothing
End Function

Public Function ExecParamQuery(ByVal SQL As String, _
                               ParamArray Params()) As ADODB.Recordset
  Dim mRes As ADODB.Recordset
  Dim mParamArr As Variant, mParam As Variant
  Dim i As Long
  
  Set m_Command = New ADODB.Command

  Call OpenConn
  m_Command.ActiveConnection = m_Conn
  m_Command.CommandText = SQL
  m_Command.CommandType = adCmdText
  
  mParamArr = Params
  If VarType(Params(0)) = 8204 Then
    mParamArr = Params(0)
  End If

  With m_Command
    For Each mParam In mParamArr
      Dim Para As ADODB.Parameter

      Set Para = .CreateParameter(CStr(i), m_GetVarType(mParam), adParamInput, LenB(mParam))
      Para.Value = mParam
      If VarType(mParam) = vbEmpty Then
        Para.Size = 1
      ElseIf VarType(mParam) = vbString Then
        If LenB(mParam) = 0 Then
          Para.Size = 1
        End If
      End If
      .Parameters.Append Para
    Next
  End With
  
  Set mRes = m_Command.Execute()
  mRes.ActiveConnection = Nothing
  Call CloseConn
  Set ExecParamQuery = mRes
  Set m_Command = Nothing
End Function

Public Function ExecNamedQuery(ByVal SQL As String, HashedParams As CHashTable) As ADODB.Recordset
  Dim mParams() As Variant
  m_ProcessNameParams SQL, HashedParams, mParams
  Set ExecNamedQuery = ExecParamQuery(SQL, mParams)
End Function

Public Function ExecNonQuery(ByVal SQL As String) As Long
  Dim affectedRows As Long

  Set m_Command = New ADODB.Command
  Call OpenConn
  m_Command.ActiveConnection = m_Conn
  m_Command.CommandText = SQL
  m_Command.CommandType = adCmdText
  m_Command.Execute affectedRows
  Call CloseConn
  Set m_Command = Nothing
  ExecNonQuery = affectedRows
End Function

Public Function ExecParamNonQuery(ByVal SQL As String, ParamArray Params()) As Long
  Dim i As Long, affectedRows As Long
  Dim mParamArr As Variant, mParam As Variant

  Set m_Command = New ADODB.Command
  Call OpenConn
  m_Command.ActiveConnection = m_Conn
  m_Command.CommandText = SQL
  m_Command.CommandType = adCmdText
  mParamArr = Params
  If VarType(Params(0)) = 8204 Then
    mParamArr = Params(0)
  End If
  With m_Command
    For Each mParam In mParamArr
      Dim Para As ADODB.Parameter

      Set Para = .CreateParameter(CStr(i), m_GetVarType(mParam), adParamInput, LenB(mParam))
      Para.Value = mParam
      If VarType(mParam) = vbEmpty Then
        Para.Size = 1
      ElseIf VarType(mParam) = vbString Then
        If LenB(mParam) = 0 Then
          Para.Size = 1
        End If
      End If
      .Parameters.Append Para
    Next
  End With

  m_Command.Execute affectedRows
  Call CloseConn
  Set m_Command = Nothing
  ExecParamNonQuery = affectedRows
End Function

Public Function ExecNamedNonQuery(ByVal SQL As String, HashedParams As CHashTable) As Long
  Dim mParams() As Variant
  m_ProcessNameParams SQL, HashedParams, mParams
  ExecNamedNonQuery = ExecParamNonQuery(SQL, mParams)
End Function

Public Function ExecCreate(ByVal SQL As String) As Variant
  Dim mRes As ADODB.Recordset

  Set m_Command = New ADODB.Command
  Call OpenConn
  
  m_Command.ActiveConnection = m_Conn
  m_Command.CommandText = SQL
  m_Command.CommandType = adCmdText
  m_Command.Execute
  
  m_Command.CommandText = "SELECT @@identity"
  Set mRes = m_Command.Execute
  If mRes.RecordCount > 0 Then
    ExecCreate = mRes.Fields(0).Value
  Else
    ExecCreate = Empty
  End If
  Call CloseConn
  Set m_Command = Nothing
  Set mRes = Nothing
End Function

Public Function ExecParamCreate(ByVal SQL As String, ParamArray Params()) As Variant
  Dim mParamArr As Variant, mParam As Variant
  Dim mRes As ADODB.Recordset
  Dim i As Long

  Set m_Command = New ADODB.Command
  Call OpenConn
  m_Command.ActiveConnection = m_Conn
  m_Command.CommandText = SQL
  m_Command.CommandType = adCmdText
  mParamArr = Params
  If VarType(Params(0)) = 8204 Then
    mParamArr = Params(0)
  End If
  With m_Command
    For Each mParam In mParamArr
      Dim Para As ADODB.Parameter

      Set Para = .CreateParameter(CStr(i), m_GetVarType(mParam), adParamInput, LenB(mParam))
      Para.Value = mParam
      If VarType(mParam) = vbEmpty Then
        Para.Size = 1
      ElseIf VarType(mParam) = vbString Then
        If LenB(mParam) = 0 Then
          Para.Size = 1
        End If
      End If
      .Parameters.Append Para
    Next
  End With

  m_Command.Execute
  m_Command.CommandText = "SELECT @@identity"
  Set mRes = m_Command.Execute
  If mRes.RecordCount > 0 Then
    ExecParamCreate = mRes.Fields(0).Value
  Else
    ExecParamCreate = Empty
  End If
  Call CloseConn
  Set m_Command = Nothing
  Set mRes = Nothing
End Function

Public Function ExecNamedCreate(ByVal SQL As String, HashedParams As CHashTable) As Variant
  Dim mParams() As Variant
  m_ProcessNameParams SQL, HashedParams, mParams
  ExecNamedCreate = ExecParamCreate(SQL, mParams)
End Function

Public Function ExecScalar(ByVal SQL As String) As Variant
  Dim mRes As ADODB.Recordset
  
  Set mRes = ExecQuery(SQL)
  If mRes.RecordCount <= 0 Then
    ExecScalar = Empty
  Else
    ExecScalar = mRes.Fields(0).Value
  End If
  Set mRes = Nothing
End Function

Public Function ExecParamScalar(ByVal SQL As String, _
                                     ParamArray Params()) As Variant
  Dim mRes As ADODB.Recordset
  
  If VarType(Params(0)) = 8204 Then
    Params = Params(0)
  End If
  
  Set mRes = ExecParamQuery(SQL, Params)
  If mRes.RecordCount <= 0 Then
    Set ExecParamScalar = Nothing
  Else
    ExecParamScalar = mRes.Fields(0).Value
  End If
  Set mRes = Nothing
End Function

Public Function ExecNamedScalar(ByVal SQL As String, HashedParams As CHashTable) As Variant
  Dim mParams() As Variant
  m_ProcessNameParams SQL, HashedParams, mParams
  ExecNamedScalar = ExecParamScalar(SQL, mParams)
End Function
