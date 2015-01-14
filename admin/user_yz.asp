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
<%Dim id,str,i,qxyz,page
Server.ScriptTimeout=50000
page=request("page")
If page="" Then page=Session("page")
id=trim(request("selectedid"))
qxyz=trim(request("qxyz"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('没有选择记录！');" &"window.location='users_List.asp'" & "</script>"
response.end
end If
Call OpenConn
str=split(id,",")
for i=0 to ubound(str)
set rs=server.createobject("adodb.recordset")
sql="select useryz,mailyz from [DNJY_user] where username='"&trim(str(i))&"'"
rs.open sql,conn,1,3
if qxyz="uno" then
rs("useryz")=0
elseif qxyz="uyes" then
rs("useryz")=1
elseif qxyz="mno" then
rs("mailyz")=0
elseif qxyz="myes" then
rs("mailyz")=1
End If
rs.update
Next
Call CloseRs(rs)
Call CloseDB(conn)
Session("page")=""
response.Write "<script language=javascript>location.replace('users_List.asp?page="&page&"');</script>"%>