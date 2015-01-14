<%@language=vbscript codepage=936 %>
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
Response.cookies("administrator")=""
Response.cookies("admindick")=""
session("admin_dick")=""
Response.Redirect "index.asp"
%>
