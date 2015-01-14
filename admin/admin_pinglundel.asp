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
dim id,str2,i
Server.ScriptTimeout=50000
id=trim(request("ID"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('没有选择记录！');" &"window.location='admin_comment.asp'" & "</script>"
response.end
end if
str2=split(id,",")
for i=0 to ubound(str2)'循环开始
set rs=server.createobject("adodb.recordset")
rs.open "delete from DNJY_news_pinglun where pinglunid="&cstr(str2(i)),conn,3
set rs=Nothing
Next'循环结束
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('成功删除！');" & chr(13)
        If request.querystring("pl")="yes" Then
        response.write "window.document.location.href='admin_Comment.asp';"&chr(13)
		else
		response.write "window.document.location.href='admin_article.asp';"&chr(13)
		End if
		response.write "</script>" & chr(13)
response.end
Call CloseDB(conn)%>