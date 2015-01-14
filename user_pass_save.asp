<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="config.asp"-->
<!--#include file=usercookies.asp-->
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
<%dim username,pass1,pass2,pass3,pass_1,pass_2,problem,answer1,answer2,answer3,answer_1,answer_2
username=request.cookies("dick")("username")
pass1=trim(request("pass1"))
pass2=trim(request("pass2"))
pass3=trim(request("pass3"))
pass_1=md5(pass1)
pass_2=md5(pass2)
problem=trim(request("problem"))
answer1=trim(request("answer1"))
answer2=trim(request("answer2"))
answer3=trim(request("answer3"))
answer_1=md5(answer1)
answer_2=md5(answer2)

Call OpenConn

If request("key")="pass" then
set rs=server.createobject("adodb.recordset")
sql="select password from [DNJY_user] where username='"&username&"' and password='"&pass_1&"'"
rs.open sql,conn,1,3
if rs.eof then
response.write "<script language=JavaScript>" & chr(13) & "alert('原密码错误！');" & "history.back()" & "</script>"
else
rs("password")=pass_2
rs.update
response.Write "<script language=javascript>alert('修改密码成功，请保管好你的新密码!');location.replace('user_pass.asp');</script>"
end If
Call CloseRs(rs)
End If

If request("key")="answer" Then
set rs=server.createobject("adodb.recordset")
If request("Modify")="ok" then
sql="select problem,answer from [DNJY_user] where username='"&username&"' and answer='"&answer_1&"'"
rs.open sql,conn,1,3
if rs.eof then
response.write "<script language=JavaScript>" & chr(13) & "alert('原密码保护答案错误！');" & "history.back()" & "</script>"
Else
rs("problem")=problem
rs("answer")=answer_2
rs.update
response.Write "<script language=javascript>alert('修改密码保护成功，请保管好你的新密码保护答案!');location.replace('user_pass.asp');</script>"
end If
Else
sql="select problem,answer from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,3
rs("problem")=problem
rs("answer")=answer_2
rs.update
response.Write "<script language=javascript>alert('密码保护设置成功，请保管好你的密码保护资料!');location.replace('user_pass.asp');</script>"
end If
Call CloseRs(rs)
End If

Call CloseDB(conn)
%>
