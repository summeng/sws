<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%dim referer
referer = Request.ServerVariables("HTTP_REFERER")
Call OpenConn
if Request("action") = "headyes" then
 call headyes()
elseif request("action") = "headno" then
 call headno()
elseif request("action") = "tuijianno" then
 call tuijianno()
elseif request("action") = "tuijianyes" then
 call tuijianyes()
elseif request("action") = "ifshowno" then
 call ifshowno()
elseif request("action") = "ifshowyes" then
 call ifshowyes()

%>


<%sub headno()
set rs=server.createobject("adodb.recordset")
sql="select id,newstop from DNJY_News where id="&request("id")
rs.open sql,conn,1,3
rs("newstop")=0
rs.update
Call CloseRs(rs)
response.redirect (referer)
end sub%>


<%sub headyes()
set rs=server.createobject("adodb.recordset")
sql="select id,newstop from DNJY_News where id="&request("id")
rs.open sql,conn,1,3
rs("newstop")=1
rs.update
Call CloseRs(rs)
response.redirect (referer)
end sub%>


<%sub tuijianno()
set rs=server.createobject("adodb.recordset")
sql="select id,tuijian from DNJY_News where id="&request("id")
rs.open sql,conn,1,3
rs("tuijian")=0
rs.update
Call CloseRs(rs)
response.redirect (referer)
end sub%>


<%sub tuijianyes()
set rs=server.createobject("adodb.recordset")
sql="select id,tuijian from DNJY_News where id="&request("id")
rs.open sql,conn,1,3
rs("tuijian")=1
rs.update
Call CloseRs(rs)
response.redirect (referer)
end sub%>


<%sub ifshowno()
set rs=server.createobject("adodb.recordset")
sql="select id,ifshow from DNJY_News where id="&request("id")
rs.open sql,conn,1,3
rs("ifshow")=0
rs.update
Call CloseRs(rs)
response.redirect (referer)
end sub%>


<%sub ifshowyes()
set rs=server.createobject("adodb.recordset")
sql="select id,ifshow from DNJY_News where id="&request("id")
rs.open sql,conn,1,3
rs("ifshow")=1
rs.update
Call CloseRs(rs)
response.redirect (referer)
end sub

else
response.write"<script language=javascript>alert('对不起，你没有此管理权限！');"
response.write"javascript:history.go(-1)</script>"
response.end
end if
Call CloseDB(conn)%>

