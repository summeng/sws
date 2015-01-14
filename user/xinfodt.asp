<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<%dim id,ServerURL,n
id=strint(request.querystring("id"))
ServerURL=CStr(Request.ServerVariables("SCRIPT_NAME"))
n=InStrRev(ServerURL,"/") '从右边第一个字符起查找"_"的位置,n为返回值
ServerURL=left(ServerURL, n)'显示从左边数第"n"个字符前面的字符,
ServerURL="http://"&Request.ServerVariables("SERVER_NAME")&""&ServerURL
Call OpenConn
if request("action")="Bad" Then'显示发布时间
Dim rs9
set rs9=server.createobject("adodb.recordset")
sql = "select * from DNJY_Bad_info where sh=1 and infoid="&cstr(id)
rs9.open sql,conn,1,3
response.Write "document.write('"&rs9.recordcount&"');" & vbCrLf
rs9.close
set rs9=Nothing
end If
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_info where yz=1 and id="&cstr(id)
if jiaoyi=0 then
sql=sql&" and zt="&jiaoyi
end If
if overdue=0 Then
sql=sql&" and dqsj>"&DN_Now&""
end If
sql=sql&" and id="&cstr(id)
rs.open sql,conn,1,3
if request("action")="dqsj" then
Dim sj,fbsj
fbsj=rs("fbsj")
if rs("dqsj")<>"" Then
sj=DateDiff("d",now(),""&rs("dqsj")&"")
if sj>0 then
response.Write "document.write('<font color=#FF0000> 剩余"&sj&"</b>天</font>');" & vbCrLf
elseif sj=0 then
response.Write "document.write('<font color=""#414141"">今日到期</font>');" & vbCrLf
elseif sj<0 then
response.Write "document.write('<font color=""#FF0000""> 已经过期</font>');" & vbCrLf
end If
end If
end If

if request("action")="llcs" Then'显示浏览次数
response.Write "document.write('"&rs("llcs")&"');" & vbCrLf
rs("llcs")=rs("llcs")+1
rs.update
end If

if request("action")="hfcs" Then'显示回复次数
response.Write "document.write('"&rs("hfcs")&"');" & vbCrLf
end If

if request("action")="zt" Then'显示交易状态
dim zt
zt=rs("zt")
if zt=1 then
response.write "document.write('<font color=""#ff0000"">已经完成交易</font>');" & vbCrLf
elseif zt=0 then
response.write "document.write('<font color=""#0066FF"">还未完成交易</font>');" & vbCrLf
end if
end If

if request("action")="username" And request.cookies("dick")("username")<>"" Then'显示回复区域用户名
response.Write "document.write('<= <FONT color=blue><FONT style=""CURSOR: hand"" onclick=username.value="""&request.cookies("dick")("username")&""" color=#0000ff>"&request.cookies("dick")("username")&"</FONT></FONT>');" & vbCrLf
end If

if request("action")="email" And request.cookies("dick")("username")<>"" Then'显示回复区域信箱
Dim usermail
If request.cookies("dick")("username")<>"" Then usermail=conn.Execute("Select email From DNJY_user Where username='"&request.cookies("dick")("username")&"'")(0)
response.Write "document.write('<= <FONT color=blue><FONT style=""CURSOR: hand"" onclick=mymail.value="""&usermail&""" color=#0000ff>"&usermail&"</FONT></FONT>');" & vbCrLf
end If

if request("action")="usernumber" Then'显示加盟会员
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_user"
rs.open sql,conn,1,1
response.write "document.write('在线经销商"&rs.recordcount&"位');" & vbCrLf
End If

Call CloseRs(rs)
Call CloseDB(conn)%>
