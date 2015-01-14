<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<%
dim id,username,email,a,b,c,d
a=trim(request("a"))
b=trim(request("b"))
c=trim(request("c"))
d=trim(request("d"))
id=trim(request("id"))
if not isnumeric(a) or a="" then
response.write "<li>道具只能为数字！"
cl
response.end
end if
if not isnumeric(b) or b="" then
response.write "<li>道具只能为数字！"
cl
response.end
end if
if not isnumeric(c) or c="" then
response.write "<li>道具只能为数字！"
cl
response.end
end if
if not isnumeric(d) or d="" then
response.write "<li>道具只能为数字！"
cl
response.end
end if
if not isnumeric(id) or id="" then
response.write "<li>参数错误1！"
cl
response.end
end If
Call OpenConn
set rs=server.createobject("adodb.recordset")
sql="select a,b,c,d from [DNJY_user] where id="&id
rs.open sql,conn,1,3
if rs.eof or rs.bof then
response.write "<li>参数错误！"
cl
response.end
end if
rs("a")=a
rs("b")=b
rs("c")=c
rs("d")=d
rs.update
Call CloseRs(rs)
Call CloseDB(conn)
response.write "<li>修改道具成功！"
call cl()
%>
<%sub cl()%>
<body onLoad="setTimeout(window.close, 2000)">
<%end sub%>
