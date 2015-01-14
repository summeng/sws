<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("06")
Call OpenConn
dim rstp,rspl,str2,i,id,firstImageName,objFSO,rsuser
Server.ScriptTimeout=50000
id=trim(request("DelID"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('没有选择记录！');" &"window.location='admin_article.asp'" & "</script>"
response.end
end if
str2=split(id,",")
set rs=server.createobject("adodb.recordset")
set rstp = server.CreateObject ("Adodb.recordset")
set rspl = server.CreateObject ("Adodb.recordset")
for i=0 to ubound(str2)'循环开始
'===============删除相关图片===========================
set rstp = server.CreateObject ("Adodb.recordset")
sql="select firstImageName,username from [DNJY_News] where id="&cstr(str2(i))
rstp.open sql,conn,1,1
if Not IsNull(rstp("firstImageName")) Then
   if Not rstp.eof Then firstImageName=rstp("firstImageName")
   call deltu()  
end If
'===============================================
'===============扣会员积分===========================
If isnull(rstp("username")) then
set rsuser=server.createobject("adodb.recordset")
sql="select * from [DNJY_user] where username='"&rstp("username")&"'"
rsuser.open sql,conn,1,3    
if Not rsuser.eof Then
rsuser("jf")=rsuser("jf")-tgjf'扣会员积分
rsuser.update
Call CloseRs(rsuser)
end If
end if
'===============================================
'===============删除相关评论===========================
set rspl = server.CreateObject ("Adodb.recordset")
rspl.Open "select * from DNJY_news_pinglun where id="&cstr(str2(i)),conn,3,3
do while not rspl.eof
rspl.delete
rspl.movenext
Loop
'===============================================

sql="delete from [DNJY_News] where id="&cstr(str2(i))
rs.open sql,conn,1,3

rspl.close
set rspl=Nothing
rstp.close
set rstp=nothing
Next'循环结束


response.write "<script language='javascript'>" & chr(13)
		response.write "alert('成功删除！');" & chr(13)
		response.write "window.document.location.href='admin_article.asp';"&chr(13)
		response.write "</script>" & chr(13)
response.end
Call CloseRs(rs)
rspl.close
set rspl=Nothing
rstp.close
set rstp=nothing

sub deltu()
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.fileExists(Server.MapPath("../"&firstImageName)) then
objFSO.DeleteFile(Server.MapPath("../"&firstImageName))
end if
set objfso=nothing
end sub
Call CloseDB(conn)%>
