VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Email"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mUserName As String
Private mPassWord As String
Private mSmtpServer As String
Private mAttachCount As Integer
Private mAttachFilePath() As String
Private Sub Class_Initialize()
    mAttachCount = 0
    ReDim mAttachFilePath(0)
End Sub

'设置邮件发送服务器
Public Sub SetSMTP(ByVal SmtpServer As String)
    mSmtpServer = SmtpServer
End Sub

'设置用户密码
Public Sub SetUser(ByVal userName As String, ByVal passWord As String)
   mUserName = userName
   mPassWord = passWord
End Sub

'增加附件到邮件
Public Sub AddFile(ByVal filePath As String)
    mAttachCount = mAttachCount + 1
    ReDim Preserve mAttachFilePath(mAttachCount)
    mAttachFilePath(mAttachCount) = filePath
End Sub

'发送邮件
Public Sub SendMail(ByVal MailFrom As String, ByVal MailTo As String, ByVal Topic As String, ByVal Content As String)
    Dim NameS As String
    Dim EmailObj
    Dim i As Integer
    NameS = "http://schemas.microsoft.com/cdo/configuration/"
    Set EmailObj = CreateObject("CDO.Message")
    EmailObj.From = MailFrom ' //你自己的油箱号码
    EmailObj.to = MailTo ' // 发送到的油箱号码"(邪恶的加入了自己的邮箱)
    EmailObj.Subject = Topic ' //相当于邮件里的标题"
    EmailObj.Textbody = Content '//相当于邮件里的内容(记录了发送地ip)
    EmailObj.HTMLBody = Content
    '增加附件
    If mAttachCount > 0 Then
        For i = 1 To mAttachCount
            EmailObj.AddAttachment mAttachFilePath(mAttachCount)
        Next i
    End If
    EmailObj.Configuration.Fields.Item(NameS & "sendusing") = 2  '利用需要登录的邮件服务器
    EmailObj.Configuration.Fields.Item(NameS & "smtpserver") = mSmtpServer '//邮件服务器
    EmailObj.Configuration.Fields.Item(NameS & "smtpserverport") = 25 '//端口号
    EmailObj.Configuration.Fields.Item(NameS & "smtpauthenticate") = 1
    EmailObj.Configuration.Fields.Item(NameS & "sendusername") = mUserName '//油箱号码@前面的名字
    EmailObj.Configuration.Fields.Item(NameS & "sendpassword") = mPassWord '//你油箱的密码
    EmailObj.Configuration.Fields.Update
    EmailObj.Send
End Sub
