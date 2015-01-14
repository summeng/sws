<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=../Include/md5.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
 <link href="../css/style.css" rel="stylesheet" type="text/css">
<%dim username,password
If Request.Cookies("status")="" Then
    Response.Write "你浏览器Cookie没有开启，无法正常运行本程序!"
    response.end
 end If	
	if session("dick_err")="" then
	session("dick_err")=0	
	end If

If trim(request.form("admin_codee"))="" Then
response.Write "<script language=javascript>alert('请输入验证码!');location.replace('login.asp');</script>"
response.end
else
if CInt(trim(request.form("admin_codee")))<>CInt(Session("ttpSN")) then
response.Write "<script language=javascript>alert('验证码不正确!');location.replace('login.asp');</script>"
Session("ttpSN")=""
response.end
end If
end If
Call OpenConn
	username=replace(trim(request("username")),"'","‘")
	password=md5(replace(trim(request("password")),"'","‘"))

	set rs=server.createobject("adodb.recordset")
	sql="select * from [DNJY_admin] where Password='"&Password&"' and UserName='"&UserName&"'"
	rs.open sql,conn,1,1
 	if not rs.bof and not rs.eof then
 		if password=rs("password") Then			
        Response.cookies("administrator")=rs("username")
		Response.cookies("dick")("username")=rs("username")'直接登录前台
        Response.cookies("dick")("domain")=Request.ServerVariables("SERVER_NAME")'直接登录前台
        Response.cookies("admindick")=1		  
        'Response.Cookies("dick").Expires=DateAdd("n",20,now())'设置cookies("dick")20分钟内有效
        Response.cookies("domain")=Request.ServerVariables("SERVER_NAME")
		session("admin_dick")=now()

       Response.Redirect "index.asp"
 		Else
		    session("dick_err")=session("dick_err")+1
            response.write " <b>身份确认失败"&session("dick_err")&"次，连续错误超"&htdlip&"次将锁定IP，无法访问！请<a href=login.asp><b><font color=#Ff0000>重新登陆</font></b></a></b> "
            
            response.end
 		end if
	Else
	        session("dick_err")=session("dick_err")+1
            response.write " <b>身份确认失败"&session("dick_err")&"次，连续错误超"&htdlip&"次将锁定IP，无法访问！请<a href=login.asp><b><font color=#Ff0000>重新登陆</font></b></a></b> "
            
            response.end
	end if
Call CloseRs(rs)
Call CloseDB(conn)%>
 <LINK href="../css/css.css" rel=stylesheet></LINK>

