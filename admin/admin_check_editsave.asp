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
end if
id=request("id")
set rs=server.CreateObject("ADODB.RecordSet")
rs.open "select * from DNJY_wenda where id="&id,conn,1,3
rs("wenti")=wenti
rs("daan")=daan
rs.update
Call CloseRs(rs)
Call CloseDB(conn)
response.write"<SCRIPT language=JavaScript>alert('问题和答案修改成功！');"
response.write"location.href=""admin_check.asp"";</SCRIPT>"
Response.End
%>