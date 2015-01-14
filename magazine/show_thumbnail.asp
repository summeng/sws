<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Dim server_a,server_b,ErrMsg,issueid
server_a=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_b=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_a,8,len(server_b))<>server_b then
ErrMsg="禁止从站点外部提交数据！"
response.write errmsg
response.end
end if
Call OpenConn
issueid=Request.Cookies("issueid")
If issueid="" Then
response.redirect "img/ad.jpg"
response.end
End if
set rs=server.CreateObject("ADODB.Recordset")
sql="select * from DNJY_zz_edition where state="&DN_True&" and id="&issueid&""
rs.open sql,conn,1,1
if rs.eof or rs.bof Then
response.redirect "img/ad.jpg"
elseIf request("dt")="ok" Then
response.redirect "../"&rs("largerpic")&""
else
response.redirect "../"&rs("thumbnail")&""
End if
Call CloseRs(rs)
Call CloseDB(conn)
%>
