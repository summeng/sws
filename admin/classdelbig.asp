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
<%Call OpenConn
dim bigclassname
bigclassname=trim(request("bigclassname"))
if bigclassname<>"" then
	sql="delete from DNJY_bigClass where bigclassname='" & bigclassname & "'"
	conn.execute sql
	sql="delete from DNJY_SmallClass where bigclassname='" & bigclassname & "'"
	conn.execute sql
	sql="delete from DNJY_News where bigclassname='" & bigclassname & "'"
	conn.execute sql
end if
Call CloseDB(conn)    
response.redirect "classmanage.asp"
%>




                                                              
                                                              
                                                              