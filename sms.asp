<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="config.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'技术支持论坛 http://www.ip126.com/bbs
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<% response.charset="gb2312" %>
<%
response.Write "免费版没有此功能，不能使用手机短信验证。在后台网站系统管理－基本参数中关闭手机短信验证功能！"
response.end
%>