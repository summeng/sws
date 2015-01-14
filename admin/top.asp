<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../Include/Version.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=GB2312" />
<link rel="stylesheet" href="inc/style_top.css" type="text/css" />
</head>
<body>
<div id="body1">→ 欢迎</FONT>&nbsp;<b><%=request.cookies("administrator")%></b>&nbsp;进入控制面版</div>
<div id="body2" class='spa'><%Response.Write "<A HREF='http://www.ip126.com' TARGET='_blank'>TianTian &reg; INFO of SAD&#8482; " & SystemVersion & " "&DatabaseType &" Build " & SystemBuildDate & "</a>"%></div>
</body>
</html>