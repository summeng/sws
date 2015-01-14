<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="config.asp"-->
<!--#include file=Include/ipt.asp-->
<!--#include file=Include/err.asp-->
<!--#include file=Include/md5.asp-->
<!--#include file="Include/head.asp"-->
<!--#include file="Include/Function.asp"-->
<!--#include file="Include/label.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
dim username,password,k,userip,email
if trim(request("yzm"))<>Session("pSN") Then
response.Write "<script language=javascript>alert('对不起，验证码不正确!');location.replace('login.asp?"&C_ID&"');</script>"

response.end
end if
if request.cookies("dick")("errs")="" then
k=0
else
k=request.cookies("dick")("errs")
end if
if request.cookies("dick")("errs")>10000 then
errdick(8)
response.end
end if
username=replace(trim(request("username2")),"'","‘")
password=md5(request("password"))
chkdick(username)
if nothaveChinese(username)=false then
errdick(3)
end If
Call OpenConn
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_user] where UserName='"&UserName&"' and password='"&password&"'"
rs.open sql,conn,1,3
if rs.eof then
Response.cookies("dick")("errs")=k+1
session("login_error")=session("login_error")+1
response.Write "<script language=javascript>alert('用户名不存在或密码错误!');location.replace('login.asp?"&C_ID&"');</script>"
Call CloseRs(rs)
Call CloseDB(conn)
response.end
Else
If usersh=1 And rs("useryz")=0 Then'判断是否审核通过
response.Write "<script language=javascript>alert('您的会员资格等待审核中...');location.replace('index.asp?"&C_ID&"');</script>"
Call CloseRs(rs)
Call CloseDB(conn)
response.end
End if
If rs("dldata")<>"" Then Response.cookies("dick")("dldata")=rs("dldata")
Response.cookies("dick")("errs")=0
session("login_error")=0
rs("dldata")=now()
if DateDiff("h",rs("jftime"),Now())>24 then'距上次加分大于24小时则加分
rs("jf")=rs("jf")+jf_3
rs("jftime")=Now()
End if
rs("dlcs")=rs("dlcs")+1
userip=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If userip="" Then 
userip=Request.ServerVariables("REMOTE_ADDR")
end if
rs("userip")=userip
rs.update
Response.cookies("dick")("username")=rs("username")
Response.cookies("dick")("domain")=Request.ServerVariables("SERVER_NAME")
Response.cookies("dick")("id")=rs("id")
Response.cookies("dick")("email")=rs("email")

end if
Call CloseRs(rs)
Call CloseDB(conn)
Response.Redirect "index.asp?"&C_ID&""
%>
