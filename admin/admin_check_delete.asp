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
set rs=server.createobject("adodb.recordset")
rs.open "delete from DNJY_wenda where id="&request.querystring("id"),conn,3
set rs=Nothing
Call CloseDB(conn)
response.write"<SCRIPT language=JavaScript>alert('问题和答案删除成功！');"
response.write"location.href=""admin_check.asp"";</SCRIPT>"
Response.End
%>

