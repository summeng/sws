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
<%Call OpenConn
dim id,i,str2,dick,bdback
bdback = Request.ServerVariables("HTTP_REFERER")
dick=request("dick")
id=trim(request("selectedid"))
str2=split(id,",")
for i=0 to ubound(str2)
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_Bad_info] where id="&trim(str2(i))
rs.open sql,conn,1,3

if dick="delyz" then
rs("sh")=0
else
rs("sh")=1
end if
rs.update
Next
response.redirect (bdback)
Call CloseDB(conn)%>