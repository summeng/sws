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
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
dim username,password,k,userip,email
if trim(request("yzm"))<>Session("pSN") Then
response.Write "<script language=javascript>alert('�Բ�����֤�벻��ȷ!');location.replace('login.asp?"&C_ID&"');</script>"

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
username=replace(trim(request("username2")),"'","��")
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
response.Write "<script language=javascript>alert('�û��������ڻ��������!');location.replace('login.asp?"&C_ID&"');</script>"
Call CloseRs(rs)
Call CloseDB(conn)
response.end
Else
If usersh=1 And rs("useryz")=0 Then'�ж��Ƿ����ͨ��
response.Write "<script language=javascript>alert('���Ļ�Ա�ʸ�ȴ������...');location.replace('index.asp?"&C_ID&"');</script>"
Call CloseRs(rs)
Call CloseDB(conn)
response.end
End if
If rs("dldata")<>"" Then Response.cookies("dick")("dldata")=rs("dldata")
Response.cookies("dick")("errs")=0
session("login_error")=0
rs("dldata")=now()
if DateDiff("h",rs("jftime"),Now())>24 then'���ϴμӷִ���24Сʱ��ӷ�
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
