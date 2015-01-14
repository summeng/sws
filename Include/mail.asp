<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Call OpenConn
dim mailname,mailpass,mailform,mailsmtp
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from DNJY_mail order by id "&DN_OrderType&""
rs.open sql,conn,3,3
if rs.eof then
response.write "还没有邮件系统参数！请在“网站系统管理-邮件系统”中添加系统邮件发送参数！"
Call CloseRs(rs)
Call CloseDB(conn)
response.End
End if
mailname=rs("mailname") '登录用户名
mailpass=rs("mailpass") '登录密码
mailform=rs("mailform") '发信人信箱
mailsmtp=rs("mailsmtp") '邮件服务器地址
Call CloseRs(rs)
Call CloseDB(conn)

'****************************************************
'函数名：ChkMail
'作 用：邮箱格式检测
'参 数：Email ----Email地址
'返回值：True正确，False有误
'使用示例：
'If ChkMail("ls535427@2221262.com") = True Then
'Response.Write "格式正确"
'Else
'Response.Write "格式有误"
'End If
'****************************************************
Public Function ChkMail(ByVal Email)
Dim Rep,Pmail : ChkMail = True : Set Rep = New RegExp
Rep.Pattern = "([\.a-zA-Z0-9_-]){2,10}@([a-zA-Z0-9_-]){2,10}(\.([a-zA-Z0-9]){2,}){1,4}$"
Pmail = Rep.Test(Email) : Set Rep = Nothing
If Not Pmail Then ChkMail = False
End Function

'**********************************jmail发信函数************************************
'制作：天天网络科技www.ip126.com
'函数名：DNJYEmail() 
'作用：利用Jmail4.3组件发送E-Mail
'参数：
'Email:类型：字符串。作用：接收E-Mail的地址。
'HostName邮件发送者名字
'E_Subject:类型：字符串。作用：信件主题。
'Information:类型：字符串。作用：信件内容。
'S_Type:类型：布尔值。作用：是否为Html格式信件。True为Html格式。False为文本格式。 
'C_M_Chk:类型：布尔值。作用：Smtp服务器是否需要身份验证 ,True需要验证，False不需要验证。
'如果发送成功，函数将返回True否则返回False 
'***********************************************************************************

Function DNJYEmail(Email,HostName,E_Subject,Information,S_Type,C_M_Chk)'括号里面分别是：邮件接收人信箱，邮件发送者名字,邮件主题，邮件内容 ，信件格式，是否要验证
Dim C_Email,C_HostName,C_Smtp,C_M_User,C_M_Pass
C_Email=mailform'发送者的邮箱
C_HostName=HostName'发送者的名字
C_Smtp=mailsmtp'Smtp服务器地址
C_M_User=mailname'如果Smtp服务器需要验证身份，请输入用户名
C_M_Pass=mailpass'请输入密码


Dim DNJY_Jmail
Err.Clear
On Error Resume Next
If Email="" Or Information="" Or E_Subject="" Then
DNJYEmail="<br><li>检查收信邮箱、信件标题、信件内容是否为空！</li>" 
Exit Function
End If
set DNJY_Jmail=Server.CreateObject("Jmail.message")'建立发送邮件的对象
if err then 
DNJYEmail= "<br><li>服务器没有安装JMail组件，无法发送邮件！</li>" 
err.clear 
exit function 
end if 
DNJY_Jmail.silent = true '屏蔽例外错误，返回FALSE跟TRUE两值
DNJY_Jmail.Charset="gb2312" '邮件的文字编码为国标,不要此句可能出现乱码
DNJY_Jmail.Logging=True'启用邮件日志 
DNJY_Jmail.From=C_Email '发件人的E-MAIL地址
DNJY_Jmail.Fromname=C_HostName'发送者的名字
DNJY_Jmail.addrecipient Email'邮件收件人的地址
DNJY_Jmail.subject=E_Subject'信件主题
If S_Type=False Then'文本或HTML格式
DNJY_Jmail.appendtext Information
Else
DNJY_Jmail.AppendHtml Information
End If
DNJY_Jmail.maildomain=C_Smtp'Smtp服务器地址
If C_M_Chk=True Then'Smtp服务器验证用户名和密码
DNJY_Jmail.mailserverusername=C_M_User'登录邮件服务器所需的用户名
DNJY_Jmail.mailserverpassword=C_M_Pass'登录邮件服务器所需的密码
End If
DNJY_Jmail.Priority = 1'邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值
DNJY_Jmail.send(C_Smtp)'执行邮件发送（通过邮件服务器地址）
If Err.Number<>0 Then'判断是否发送成功
DNJYEmail="<li>邮件发送失败，请检查邮件发送系统参数和收信人信箱是否正确！"'失败返回
Else
DNJYEmail="<li>邮件发送成功！"'成功返回
'＝＝＝＝＝同时写邮件日志开始＝＝＝＝
Dim rslog
set rslog=server.createobject("adodb.recordset")
sql="select * from DNJY_log"
rslog.open sql,conn,1,3
rslog.addnew
rslog("username")=C_HostName
rslog("outbox")=C_Email
rslog("inbox")=email
rslog("title")=E_Subject
rslog("content")=Information
rslog("Sendtime")=now()
rslog("lock")=0
rslog.update
rslog.close
set rslog=nothing
'＝＝＝＝同时写邮件日志END＝＝＝＝	
End If
Err.Clear
Set DNJY_Jmail=nothing'消毁对象
End Function %>                    