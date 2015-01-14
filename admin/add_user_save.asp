<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=../config.asp-->
<!--#include file=../Include/md5.asp-->
<!--#include file=../Include/mail.asp-->
<!--#include file=cookies.asp-->
<!--#include file=../Include/err.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("04")
dim username,password,email,maillist,userip
username=request("username")
password=request("password")
If password<>"" Then
password=md5(password)
End If

email=request("email")
maillist=request("maillist")
userip=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If userip="" Then 
userip=Request.ServerVariables("REMOTE_ADDR")
end If
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_user] where username='"&username&"' or (email='"&email&"' and email<>'')"
rs.open sql,conn,1,3
if rs.eof or rs.bof Then
rs.addnew
rs("username")=username
rs("useryz")=1
rs("mailyz")=1
If password<>"" Then rs("password")=password
rs("email")=email
rs("maillist")=maillist
rs("zcdata")=now()
rs("jftime")=now()
rs("vipdq")=now()
rs.update

response.write "<li>添加会员成功！通知会员到前台完善有关资料。"
dim mailbody
mailbody="您在"&webname&"的注册会员信息：登陆帐号："&username&" <br>登陆地址为：<a target=_blank href="&web&">"&web&"</a><br>登陆成功后请先完善你的资料，否则不能正常发布信息！<br><br>【本邮件由系统自动发送，请勿回复邮件】"
'邮件发送

Dim HostName,E_Subject,Information,S_Type,C_M_Chk,e_info
Email=email
HostName=webname
E_Subject="您在"&webname&"的注册信息"
Information=mailbody
S_Type=True
C_M_Chk=True
e_info=DNJYEmail(Email,HostName,E_Subject,Information,S_Type,C_M_Chk)
response.write "<br>发送邮件检查:"&e_info&""
Call CloseDB(conn)
response.end

'邮件发送END
Else
response.write "<li>用户名或邮箱已存在，请另选！"
end If
Call CloseDB(conn)%>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
-->
</style>