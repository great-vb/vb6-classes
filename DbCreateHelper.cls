VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DbCreateHelper"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//////////////////////////////////////////////////////////////////////////////
'@@summary:请引用“Microsoft ADO Ext. 6.0 for DDL and Security”
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
'       公有变量
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
Private m_ConnStr As String

Private m_CT As New ADOX.Catalog

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

End Sub


'------------------------------------------------------------------------------
'       销毁
'------------------------------------------------------------------------------

Private Sub Class_Terminate()
  Set m_CT = Nothing
End Sub

'//////////////////////////////////////////////////////////////////////////////
'//
'//      私有方法
'//
'//////////////////////////////////////////////////////////////////////////////

Private Function m_TableExist(ByVal TableName As String) As Boolean
  Dim i As Integer
  m_TableExist = False
  If m_CT.Tables.Count > 0 Then
    For i = 0 To m_CT.Tables.Count - 1
      If LCase(m_CT.Tables(i).Name) = LCase(TableName) Then
        m_TableExist = True
        Exit Function
      End If
    Next i
  End If
End Function

Private Function m_FieldType(ByVal FieldType As String) As DataTypeEnum
  Select Case LCase(FieldType)
    Case "string"
      m_FieldType = adVarWChar
    Case "text"
      m_FieldType = adLongVarWChar
    Case "date"
      m_FieldType = adDate
    Case "currency"
      m_FieldType = adCurrency
    Case "boolean"
      m_FieldType = adBoolean
    Case "double"
      m_FieldType = adDouble
    Case "integer"
      m_FieldType = adInteger
    Case "guid"
      m_FieldType = adGUID
    Case "single"
      m_FieldType = adSingle
    Case "longbinary"
      m_FieldType = adLongVarBinary
    Case "byte"
      m_FieldType = adUnsignedTinyInt
    Case "short"
      m_FieldType = adSmallInt
    Case Else
      m_FieldType = adVarWChar
  End Select
End Function

'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有方法
'//
'//////////////////////////////////////////////////////////////////////////////

Public Sub SetDbFile(ByVal FilePath As String)
  m_ConnStr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & FilePath & ";"
  If Dir(FilePath) = "" Then
    m_CT.Create m_ConnStr
  End If
  m_CT.ActiveConnection = m_ConnStr
End Sub

Public Sub CreateTable(ByVal TableName As String, ByVal FieldNames As String)
  Dim tb As New ADOX.Table
  Dim fields() As String
  Dim field As String
  Dim i As Integer
  Dim col As ADOX.Column

  If Len(m_ConnStr) <= 0 Then
    Err.Raise 105, , "[SunSoft]未指定数据库连接，请检查"
  End If

  If Not m_TableExist(TableName) Then
    tb.Name = TableName
    fields = Split(FieldNames, ",")
    For i = 0 To UBound(fields)
      Set col = New ADOX.Column
      col.Name = Split(fields(i), ":")(0)
      col.Type = m_FieldType(Split(fields(i), ":")(1))
      If Split(fields(i), ":")(1) = "boolean" Then
        col.Attributes = ADOX.ColumnAttributesEnum.adColFixed
      Else
        col.Attributes = ADOX.ColumnAttributesEnum.adColNullable
      End If
      
      'tb.Columns.Append Split(fields(i), ":")(0), m_FieldType(Split(fields(i), ":")(1))
      tb.Columns.Append col
    Next i
    'tb.Columns.Append "adBinary", adBinary
    'tb.Columns.Append "adBoolean", adBoolean
    'tb.Columns.Append "adCurrency", adCurrency
    'tb.Columns.Append "adDate", adDate
    'tb.Columns.Append "adDouble", adDouble
    'tb.Columns.Append "adGUID", adGUID
    'tb.Columns.Append "adInteger", adInteger
    'tb.Columns.Append "adLongVarBinary", adLongVarBinary
    'tb.Columns.Append "adLongVarWChar", adLongVarWChar
    'tb.Columns.Append "adSingle", adSingle
    'tb.Columns.Append "adSmallInt", adSmallInt
    'tb.Columns.Append "adUnsignedTinyInt", adUnsignedTinyInt
    'tb.Columns.Append "adVarBinary", adVarBinary
    'tb.Columns.Append "adVarWChar", adVarWChar
    'tb.Columns.Append "adWChar", adWChar
    m_CT.Tables.Append tb
  End If

End Sub

Public Sub InitDbFromModels(ParamArray dbModels())
  Dim model As DBModel
  Dim i As Integer
  For i = LBound(dbModels) To UBound(dbModels)
    Set model = dbModels(i)
    CreateTable model.TableName, model.TableFields
  Next
End Sub
