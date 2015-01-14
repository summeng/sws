<!--#include file="../conn.asp"-->
<!--#include file=../config.asp-->
<!--#include file=../Include/mail.asp-->

<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<html>
<head>
<title>发送电子邮件</title>
<meta http-equiv="目录类型" content="文本/html; 字符集=gb2312">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>

</head>
<BODY>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
.style1 {
	color: #FFFFFF;
	font-weight: bold;
	font-size: 12px;
}
.style2 {font-size: 12px}
-->
</style>
<center>
<%Dim aJMail
Set aJMail=Server.CreateObject("JMail.Message")   
If aJMail Is Nothing Then  
response.write "<script LANGUAGE='javascript'>alert('温馨小提示：服务器不支持JMail.Message组件，请立即安装此组件，此组件用于邮件发送!');history.go(-1);</script>"
Response.End
End If
If Request("mailto")<>"" And Request.form("mymail") = "" Then  
response.write "<script LANGUAGE='javascript'>alert('温馨小提示：请填写您的邮箱，方便对方联系您!');history.go(-1);</script>"
Response.End
End If
Call OpenConn
If trim(Request("email"))="" then
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select email from DNJY_user Where username='"&trim(Request("username"))&"'"
rs.open sql,conn,1,1
if rs.eof then
response.write "未查到此用户邮箱！"
Call CloseRs(rs)
Call CloseDB(conn)
response.End
Else
email=rs("email")
End If
else
email=trim(Request("email"))
End If
mailbody=Request("mailbody")+""&email&""
If Request("Mobile")<>"" Then mailbody=mailbody+"和手机："&Request("Mobile")&""
dim userid,rs1,tomail,mailsubject,mailbody,host,ErrStr,jmail,sitename,email,i
if Request.form("action")="send" then
call mailto()
Else
call mailedit()
End if
sub mailto()
If ChkMail(email) =False Then
response.Write "<script language=javascript>alert('收信人邮箱格式错误!');window.close();</script>"
response.end
End If
If ChkMail(Request("mymail")) =False Then
response.Write "<script language=javascript>alert('您的邮箱格式错误!');window.close();</script>"
response.end
End If
if Request("mailsubject") <> "" then 	
	mailsubject = Request.form("mailsubject") '信件主题
	mailbody = mailbody+"－－－－本邮件由["&webname&"-http://"&web&"]系统自动发送，请勿直接回复。可联系邮件操作人邮箱："&Request("mymail")&"" '信件内容
	host = ""	
	ErrStr = "ture"	

session("mcs")=session("mcs")+1
If session("mcs")>3 Then
response.Write "<script language=javascript>alert('您发信次数太密了，请休息一下!');window.close();</script>"
response.end
End If

'jmail邮件组件发送邮件开始，下面的值要先取得
Set jmail = Server.CreateObject("JMail.Message") 
jmail.silent = true '屏蔽例外错误，返回FALSE跟TRUE两值
jmail.Logging = true '启用邮件日志
jmail.Charset = "gb2312"  '邮件的文字编码为国标
If Request("radiobutton") = "html" Then
jmail.ContentType = "text/html"    '邮件的格式为HTML格式
end if
jmail.MailServerUserName = mailname '输入smtp服务器验证登陆名 （邮局中任何一个用户的Email地址） 
jmail.MailServerPassword = mailpass '输入smtp服务器验证密码 （用户Email帐号对应的密码） 
jmail.From = mailform '发信人信箱，因服务器反垃圾邮件，此信箱必须与发信服务器同域用户的信箱！ 
jmail.FromName = webname '发件人姓名 
jmail.AddRecipient email '收件人Email 
jmail.Subject = mailsubject '邮件的标题 
jmail.Body = mailbody '邮件的内容
JMail.Priority = 5      '邮件的紧急程序 
jmail.Send (mailsmtp) '执行邮件发送smtp服务器地址（企业邮局地址）
ErrStr =jmail.errormessage
jmail.close 
set jmail = nothing '关闭对象
'邮件发送结束

if ErrStr = "" then
	Response.Write "恭喜！邮件发送成功！"
'＝＝＝＝＝同时写邮件日志开始＝＝＝＝
Dim rslog
set rslog=server.createobject("adodb.recordset")
sql="select * from DNJY_log"
rslog.open sql,conn,1,3
rslog.addnew
rslog("username")=request.cookies("dick")("username")
rslog("outbox")=mailform
rslog("inbox")=email
rslog("title")=mailsubject
rslog("content")=mailbody
rslog("Sendtime")=now()
rslog("lock")=0
rslog.update
rslog.close
set rslog=nothing
'＝＝＝＝同时写邮件日志END＝＝＝＝	
else	
	Response.Write "很遗憾！邮件未能发送！<br>请检查邮件系统各项设置、是否有订阅杂志会员."
end if	'发送结果判断结束%>
<body onLoad="setTimeout(window.close, 2000)">
<%Response.end
else		     
response.write "<script language='javascript'>"
response.write "alert('出错了，邮件主题必须完整填写。');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end                                                                          
end if  '是否发信判断结束                                                                         
Call CloseDB(conn)
end Sub
sub mailedit()%> 
<table border="1" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%" cellpadding="0">
  <tr bgcolor="#BDBEDE">
    <td width="100%" height="28" colspan="8" align="center"><b><span class="style2">发送电子邮件</span></b></td>
  </tr>
</table>
  <form name="sendmail" action="user_mail.asp?mailto=yes" method="post" >
 <table border="1" width="99%" height="161"  background="images/greystrip.gif"  bordercolordark="#FFFFFF"  cellspacing="0">
      <tr>
    <td width="100%" colspan="2" height="20">
      <p ><font color="#ff0000"> 注意：请遵守有关法律，严禁利用本系统发送违法和垃圾信息！</font></p>
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="2" height="20" background="images/mmto.gif">
      <p align="center"></p>
    </td>
  </tr>
  <tr>
    <td width="18%" height="26">收信邮箱：</td>
    <td width="82%" height="26"> <%=email%></td><input type="hidden" name="tomail" value="<%=email%>"></td>              
  </tr>                                                                   
  <tr>                                                                   
    <td width="14%" height="26">邮件主题：</td>                                                                   
    <td width="86%" height="26">我对您的【<%=Request("mailzt")%>】很感兴趣<input type="hidden" name="mailsubject" value="我对您的【<%=Request("mailzt")%>】很感兴趣">
    </td>                                                                   
  </tr>    
  <tr>                                                                           
    <td width="14%" height="26">邮件内容：</td>                                                                          
    <td width="86%" height="26">您好，我对您在【<%=webname%>】发布的【<%=Request("mailzt")%>】很感兴趣，请联系我邮箱：<input type="text" name="mymail" style="border:1px black solid" size="25"><input type="hidden" name="mailbody" value="您好，我对您在【<%=webname%>】发布的【<%=Request("mailzt")%>】很感兴趣，请联系我邮箱："></td>                                                                          
  </tr> 
    <tr>                                                                   
    <td width="14%" height="26">您的手机：</td>                                                                   
    <td width="86%" height="26"><input type="text" name="Mobile" style="border:1px black solid" size="15">
    最好填写方便对方联系您</td>                                                                   
  </tr>                                                                        
  <tr>                                                                          
    <td width="100%" height="26" colspan="2">                                                                          
      <p align="center"><input type="submit" name="send" value="发送" style="border-style: solid; border-width: 1">                                                                           
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" name="Submit" value="取消" style="border-style: solid; border-width: 1">     <input type=hidden name="action" value="send"> 
	  <input type="hidden" name="username" value="<%=trim(Request("username"))%>">
	  <input type="hidden" name="email" value="<%=trim(Request("email"))%>">
    </td>                                                                          
  </tr>                                                                          
 </table></form> 
 <%end Sub%>
</center>                                                                              
</body>                                                                              
</html>
<!--#include virtual="/Include/bottom_superadmin.asp" -->