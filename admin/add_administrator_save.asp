<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file=../Include/md5.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("01")%>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
body,td,th {
	color: #FF0000;
}
-->
</style>
<%username=replace(trim(request("username")),"'","‘")
if request("edit")="ok" Then
Dim manage
	manage=""
	if request("manage01")="yes" then manage=manage+"|01"
	if request("manage02")="yes" then manage=manage+"|02"
	if request("manage03")="yes" then manage=manage+"|03"
	if request("manage04")="yes" then manage=manage+"|04"
	if request("manage05")="yes" then manage=manage+"|05"
	if request("manage06")="yes" then manage=manage+"|06"
	if request("manage07")="yes" then manage=manage+"|07"
	if request("manage08")="yes" then manage=manage+"|08"
	if request("manage09")="yes" then manage=manage+"|09"
	if request("manage10")="yes" then manage=manage+"|10"
	if request("manage11")="yes" then manage=manage+"|11"
	if request("manage12")="yes" then manage=manage+"|12"
	if request("manage13")="yes" then manage=manage+"|13"
	if request("manage14")="yes" then manage=manage+"|14"
	if request("manage15")="yes" then manage=manage+"|15"
Call OpenConn
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from DNJY_admin where username='"&username&"'"
	rs.open sql,conn,1,3
		rs("manage")=manage
		rs.update
	rs.close
	set rs=nothing
Call CloseDB(conn)
	response.write "<script language='javascript'>"
	response.write "alert('管理员"&username&"权限设置成功！');"
	response.write "location.href='add_administrator.asp?username="&username&"';"
	response.write "</script>"
else
dim username,password,password2
password=replace(trim(request("password")),"'","‘")
password2=replace(trim(request("password2")),"'","‘")
Call OpenConn
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select a.username,a.password,a.manage,b.username,b.password,c.cityadmin from DNJY_admin a,DNJY_user b,DNJY_city c where a.username='"&username&"' or b.username='"&username&"' or c.cityadmin='"&username&"'"
rs.open sql,conn,1,1
if not rs.eof or not rs.bof then
response.write "<script language=JavaScript>" & chr(13) & "alert('对不起，该帐号在前台会员或分站管理员或后台管理员中已经存在，请另选！');" & "window.location='add_administrator.asp'" & "</script>"
response.end
else
if username=""  then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，您输入的用户名不能为空！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	else
if password="" or len(password)<8 or len(password)>16 then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，您输入的密码不能为空！且要求长度为8-16位');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
else
if password<>password2 then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，确认密码不相同！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	Else
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from DNJY_admin"
rs.open sql,conn,1,3	
rs.addnew
rs("username")=username
rs("password")=md5(password)
rs("manage")="|||||||||||||||"
rs.update
Call CloseRs(rs)

set rs = Server.CreateObject("ADODB.RecordSet")
sql="select username,password,useryz,zcdata,vipdq from DNJY_user"
rs.open sql,conn,1,3
rs.addnew
rs("username")=username
rs("password")=md5(password)
rs("useryz")=1
rs("zcdata")=now()
rs("vipdq")=now()
rs.update
Call CloseRs(rs)

end if
end if
end if
end if
Call CloseRs(rs)

Call CloseDB(conn)
response.write "<script language=JavaScript>" & chr(13) & "alert('恭喜你，添加帐号成功！要完善有关资料方便发布和管理信息\n\n接下来，请为该管理员设置权限或完善有关资料！');" & "window.location='add_administrator.asp?username="&Request("username")&"'" & "</script>"
response.write "<meta http-equiv=refresh content=""1;URL=add_administrator.asp"">"
end if%>