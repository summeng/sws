<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="SqlIn.Asp"-->
<%'判断注册用户是否重复信息反馈
Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "No-Cache"
response.charset="gb2312"
name=request.querystring("name")
%>
<%Dim name,rst
If name="yes" Then
response.write("<img src='/"&strInstallDir&"images/reg_error.gif'><FONT COLOR=#ff0000>很抱歉，帐号名仅支持大小写字母、数字和下划线字符！</font>")
Else
Call OpenConn
set rs=server.createobject("adodb.recordset")
rs.open "select * from [DNJY_user] where username='"&name&"' ",conn,1,3
set rst=server.createobject("adodb.recordset")
rst.open "select * from [DNJY_usertemp] where username='"&name&"' ",conn,1,3
if (rs.eof and rs.bof) and (rst.eof and rst.bof) Then   	
response.write("<img src='/"&strInstallDir&"images/reg_ok.gif'> <FONT COLOR=#0000ff>恭喜您，此用户名可以注册！</font>")	
Else
response.write("<img src='/"&strInstallDir&"images/reg_error.gif'><FONT COLOR=#ff0000>很抱歉，此用户名已有人使用，请另选！</font>")    
end If
Call CloseRs(rs)
rst.close
set rst=Nothing
End if
Call CloseDB(conn)
%>