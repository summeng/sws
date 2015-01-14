<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file=Function.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%Call OpenConn
dim username,rsadId
username=trim(request("username"))
if Not nothaveChinese(UserName) Then
response.Write "<font color=#FF0000>本站限制不能用汉字注册!</font>"
Elseif Check_Str(UserName) then
response.Write "<font color=#FF0000>你想要注册的用户名"&username&"含有本站禁止注册的字符!</font>"
Else
set rsadId=server.CreateObject("adodb.recordset")
rsadId.open"select username from DNJY_user where  username='"&username&"'",conn,1,1
if not rsadId.eof then
response.Write "<font color=#FF0000>抱歉，用户名"&username&"已存在，请另选。</font>"
else
response.Write "<font color=#0000ff>恭喜！用户名"&username&"不存在，可以注册。</font>"
end If
rsadId.close
set rsadId=Nothing
End If
Call CloseDB(conn)%>