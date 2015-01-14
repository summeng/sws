<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
if Request.ServerVariables("SERVER_NAME")<>request.cookies("dick")("domain") then
response.redirect "login.asp?"&C_ID&""
end if
if request.cookies("dick")("username")="" or request.cookies("dick")("domain")="" then 
response.redirect "login.asp?"&C_ID&""
end if
%>
