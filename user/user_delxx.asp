<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=../config.ASP-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="../Include/md5.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
dim objFSO,fileExt,sql1,rs1,sql2,delpas
id=request("id")
if request("key")="del" then
%>
<form method="POST" action="?key=delok&id=<%=id%>">
<p align="center">请输入删除密码：<input type="text" name="delpas" size="15"> </p>
<p align="center"> <input type="submit" value="确定删除" name="B1"></p>
</form>
游客删除自己信息通道，若本信息不是您发布的请离开！
<%end if

if request("key")="delok" then
id=trim(request("id"))
delpas=md5(request("delpas"))
Call OpenConn
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_info] where delpas='"&delpas&"' and id="&cstr(id)
rs.open sql,conn,1,1
 if not(rs.eof or rs.bof) then
'================删除已生成的htm文件
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.fileExists(Server.MapPath("../html/"&trim(id)&".html")) then
objFSO.DeleteFile(Server.MapPath("../html/"&trim(id)&".html"))
end if
set objfso=Nothing
'==================================================================
set rs=server.createobject("adodb.recordset")
sql="delete from [DNJY_info] where id="&cstr(id)'删除信息
rs.open sql,conn,1,3
sql1="delete from [DNJY_favorites] where scid='"&cstr(id)&"'"'删除收藏
rs.open sql1,conn,1,3
sql2="delete from [DNJY_reply] where xxid="&cstr(id)&""'删除回复
rs.open sql2,conn,1,3
response.Write "<script language=javascript>alert('删除信息成功!');window.close();</script>"
response.end
Else
response.write "<script language=JavaScript>" & chr(13) & "alert('删除密码不正确！');" & "history.back()" & "</script>"
end If
end If
%>
