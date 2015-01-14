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
<%call checkmanage("14")
Call OpenConn
dim wenti,daan,id
wenti=request.form("wenti")
daan=request.form("daan")
if wenti="" or daan="" then
response.write"<SCRIPT language=JavaScript>alert('问题和答案都不能为空！');"
response.write"javascript:history.go(-1)</SCRIPT>"
Response.End
end If
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from DNJY_wenda"
	rs.open sql,conn,1,2

rs.addnew
rs("wenti")=wenti
rs("daan")=daan
rs("xs")=1
rs.update
Call CloseRs(rs)
Call CloseDB(conn)
response.write"<SCRIPT language=JavaScript>alert('问题和答案添加成功！');"
response.write"location.href=""admin_check.asp"";</SCRIPT>"
Response.End
%>