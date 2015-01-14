<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
if request.cookies("administrator")="" or request.cookies("admindick")<>"1" or request.cookies("domain")="" then 
response.redirect "../login.asp" 
end if
if Request.ServerVariables("SERVER_NAME")<>request.cookies("domain") then
response.redirect "../login.asp"
end If

sub checkmanage(str)'权限判断
Call OpenConn
Dim mrs,manage
Set mrs = conn.Execute("select * from DNJY_admin where username='"&request.cookies("administrator")&"'")
if not (mrs.bof and mrs.eof) then
manage=mrs("manage")
if instr(manage,str)<=0 then
conn.close
set conn = nothing
Response.Write ("<script language=javascript>alert('抱歉：你没有此项操作的权限！');history.go(-1);</script>")
response.end
end if
else 
conn.close
set conn = nothing
Response.Write ("<script language=javascript>alert('抱歉：你没有登陆，不能执行此操作！');history.go(-1);</script>")

response.end
end if
set mrs=Nothing
end sub%>
