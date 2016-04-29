VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdodbHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'////////说明////////
'使用前，请先引用：Microsoft ActiveX Data Objects 2.5 Library
'引用方式：工程→引用→Microsoft ActiveX Data Objects 2.5 Library


'最后修改日期：2015年1月15日22:16:43
'2015年01月15日22:19:23   全部重新格式化
'2016年02月09日15:33:17   更新部分Adodb缺失
'成员变量：连接对象
Private m_Conn       As ADODB.Connection

'成员变量：SQL命令对象
Private m_Command    As ADODB.Command

'成员变量：连接字符串
Private m_ConnString As String

'成员变量：数据库文件路径
Private m_FilePath   As String

'成员变量：自身对象
Private m_Me         As New AdodbHelper

'规定：每个函数执行完成后，都必须清空m_Command，并且创建一个空的对象
'【类初始化、释放】
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       Class_Initialize
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:56
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Class_Initialize()
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       Class_Terminate
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:56
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Class_Terminate()
    Set m_Conn = Nothing
    Set m_Command = Nothing
End Sub

'【类属性】
'读取：数据库连接字符串
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       ConnectionString
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:56
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get ConnectionString() As String
    ConnectionString = m_ConnString
End Property

'设置：连接字符串
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       ConnectionString
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:56
'
' Parameters :       vNewValue (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Let ConnectionString(ByVal vNewValue As String)
    m_ConnString = vNewValue
End Property

'类是否准备好，是否可以执行SQL
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       IsReady
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:56
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get IsReady() As Boolean
    IsReady = IIf(Len(ConnectionString) > 0, True, False)
End Property

'【类方法】
'执行查询，返回离线记录集
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       ExecQuery
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       SqlStr (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ExecQuery(ByVal SqlStr As String) As ADODB.Recordset

    Dim tempRes As New ADODB.Recordset

    Set m_Command = New ADODB.Command
    Call OpenConn
    m_Command.ActiveConnection = m_Conn
    m_Command.CommandText = SqlStr
    Set tempRes = m_Command.Execute()
    tempRes.ActiveConnection = Nothing
    Call CloseConn
    Set ExecQuery = tempRes
    Set m_Command = Nothing
End Function

'执行参数化查询，返回离线记录集
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       ExecParamQuery
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       SqlStr (String)
'                    Params() (Variant)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ExecParamQuery(ByVal SqlStr As String, _
                               ParamArray Params()) As ADODB.Recordset

    Dim tempRes As New ADODB.Recordset

    Dim i       As Long
    
    Dim paramArr As Variant

    Set m_Command = New ADODB.Command
    '打开连接
    Call OpenConn
    m_Command.ActiveConnection = m_Conn
    m_Command.CommandText = SqlStr
    m_Command.CommandType = adCmdText
    
    '过滤嵌套情况
    paramArr = Params
    If VarType(Params(0)) = 8204 Then
      paramArr = Params(0)
    End If

    '设置参数
    With m_Command
    

        For Each param In paramArr

            Dim Para As ADODB.Parameter

            Set Para = .CreateParameter(CStr(i), GetVarType(param), adParamInput, LenB(param))
            Para.Value = param
            If VarType(param) = vbEmpty Then
              Para.Size = 1
            ElseIf VarType(param) = vbString Then
              If LenB(param) = 0 Then
                Para.Size = 1
              End If
            End If
            .Parameters.Append Para
        Next

    End With

    '获取执行后记录集
    Set tempRes = m_Command.Execute()
    '与数据库连接脱钩
    tempRes.ActiveConnection = Nothing
    '关闭数据库连接
    Call CloseConn
    '返回数据集对象引用
    Set ExecParamQuery = tempRes
    '清空命令对象
    Set m_Command = Nothing
End Function

'执行查询，返回影响行数
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       ExecNonQuery
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       SqlStr (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ExecNonQuery(ByVal SqlStr As String) As Long

    '定义影响行数变量
    Dim affectedRows As Long

    '创建绑定Command对象
    Set m_Command = New ADODB.Command
    '打开连接
    Call OpenConn
    '绑定Command到数据库连接
    m_Command.ActiveConnection = m_Conn
    '设置SQL语句
    m_Command.CommandText = SqlStr
    '设置SQL类型
    m_Command.CommandType = adCmdText
    '获取执行后影响行数
    m_Command.Execute affectedRows
    '关闭数据库连接
    Call CloseConn
    '清空命令对象
    Set m_Command = Nothing
    '返回影响行数
    ExecNonQuery = affectedRows
End Function

'执行参数化查询，返回影响行数
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       ExecParamNonQuery
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       SqlStr (String)
'                    Params() (Variant)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ExecParamNonQuery(ByVal SqlStr As String, ParamArray Params()) As Long

    Dim i            As Long

    Dim affectedRows As Long
    
    Dim paramArr As Variant

    Set m_Command = New ADODB.Command
    '打开连接
    Call OpenConn
    m_Command.ActiveConnection = m_Conn
    m_Command.CommandText = SqlStr
    m_Command.CommandType = adCmdText
    '过滤嵌套情况
    paramArr = Params
    If VarType(Params(0)) = 8204 Then
      paramArr = Params(0)
    End If
    '设置参数
    With m_Command

        For Each param In paramArr

            Dim Para As ADODB.Parameter

            Set Para = .CreateParameter(CStr(i), GetVarType(param), adParamInput, LenB(param))
            Para.Value = param
            If VarType(param) = vbEmpty Then
              Para.Size = 1
            ElseIf VarType(param) = vbString Then
              If LenB(param) = 0 Then
                Para.Size = 1
              End If
            End If
            .Parameters.Append Para
        Next

    End With

    '获取执行后记录集
    m_Command.Execute affectedRows
    '关闭数据库连接
    Call CloseConn
    '清空命令对象
    Set m_Command = Nothing
    '返回影响行数
    ExecParamNonQuery = affectedRows
End Function

'根据文件路径设置连接字符串
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       SetConnToFile
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       FilePath (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub SetConnToFile(ByVal FilePath As String)
    m_ConnString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & FilePath & ";"
End Sub

'释放离线记录集
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       ReleaseRecordset
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       dbRes (ADODB.Recordset)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub ReleaseRecordset(ByRef dbRes As ADODB.Recordset)
    Set dbRes = Nothing
End Sub

'将制定表数据从一个数据库文件插入到另一个数据库文件中
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       TransData
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       SourceDb (String)
'                    DestDb (String)
'                    tableNames() (Variant)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub TransData(ByVal SourceDb As String, _
                     ByVal DestDb As String, _
                     ParamArray tableNames())

    Dim TableName As Variant

    m_Me.SetConnToFile SourceDb

    If UBound(tableNames) <= 0 Then

        Exit Sub

    End If

    For Each TableName In tableNames

        m_Me.ExecQuery "select * into [" & DestDb & "]." & TableName & " from " & TableName
    Next

End Sub

'从打开数据库文件，返回一个打开的数据库连接
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       DbConnFromFile
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       FilePath (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function DbConnFromFile(ByVal FilePath As String) As ADODB.Connection

    Dim tmpConn As New ADODB.Connection

    tmpConn.CursorLocation = adUseClient
    tmpConn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & FilePath & ";"
    Set DbConnFromFile = tmpConn
End Function

'执行参数化查询，返回首行，首列值
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       ExecParamQueryScalar
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       SqlStr (String)
'                    Params() (Variant)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ExecParamQueryScalar(ByVal SqlStr As String, _
                                     ParamArray Params()) As Variant

    Dim varResult As Variant

    Dim tmpRes    As ADODB.Recordset

    Set tmpRes = ExecParamQuery(SqlStr, Params)

    If tmpRes.RecordCount <= 0 Then
        Set ExecParamQueryScalar = Nothing
    Else
        ExecParamQueryScalar = tmpRes.fields(0).Value
    End If

    ReleaseRecordset tmpRes
End Function

'执行查询，返回首行，首列值
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       ExecQueryScalar
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       SqlStr (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ExecQueryScalar(ByVal SqlStr As String) As Variant

    Dim varResult As Variant

    Dim tmpRes    As ADODB.Recordset

    Set tmpRes = ExecQuery(SqlStr)

    If tmpRes.RecordCount <= 0 Then
        Set ExecQueryScalar = Nothing
    Else
        ExecQueryScalar = tmpRes.fields(0).Value
    End If

    ReleaseRecordset tmpRes
End Function

Public Function SafeParam(ByVal value As String) As String
  SafeParam = Replace(value, "'", "''")
End Function

'【内部方法】
'打开数据库连接
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       OpenConn
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub OpenConn()
    Set m_Conn = New ADODB.Connection
    m_Conn.CursorLocation = adUseClient
    m_Conn.Open ConnectionString
End Sub

'关闭数据库连接
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       CloseConn
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub CloseConn()
    m_Conn.Close
    Set m_Conn = Nothing
End Sub

'返回VB变量类型返回数据库参数的类型枚举值
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       类库构建项目
' Procedure  :       GetVarType
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       Value (Variant)
'--------------------------------------------------------------------------------
'</CSCM>
Private Function GetVarType(ByRef Value As Variant) As ADODB.DataTypeEnum

    Dim k As New ADODB.Command

    Select Case VarType(Value)

        Case VbVarType.vbString
            GetVarType = ADODB.DataTypeEnum.adVarWChar

        Case VbVarType.vbInteger
            GetVarType = ADODB.DataTypeEnum.adSmallInt

        Case VbVarType.vbBoolean
            GetVarType = ADODB.DataTypeEnum.adBoolean

        Case VbVarType.vbCurrency
            GetVarType = ADODB.DataTypeEnum.adCurrency

        Case VbVarType.vbDate
            GetVarType = ADODB.DataTypeEnum.adDate

        Case 8209
            GetVarType = ADODB.DataTypeEnum.adLongVarBinary

        Case Else
            GetVarType = ADODB.DataTypeEnum.adVariant
    End Select

End Function
