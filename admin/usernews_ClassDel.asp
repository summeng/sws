<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=../Include/Function.asp-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("04")%>

<%
dim ClassName,ClassID,sSavePathFileName,aSavePathFileName,id,i
ClassID=ReplaceUnsafe(Request.QueryString("ClassID"))
if ClassID<>"" then
	sql="delete From DNJY_userNewsClass where id=" & ClassID
	conn.Execute sql		
	
	'删除相关图片
	Set Rs = Server.CreateObject( "ADODB.Recordset" )
	Sql = "SELECT id,content,SavePathFileName from DNJY_userNews WHERE ClassID=" & ClassID
	Rs.Open Sql, Conn, 1, 3
	If Not Rs.Eof Then
		do while not Rs.EOF 
		sSavePathFileName = Rs("SavePathFileName")		
		id=Rs("id")
		if sSavePathFileName<>"" then
				' 把带"|"的字符串转为数组
			aSavePathFileName = Split(sSavePathFileName, "|")
				' 删除新闻相关的文件，从文件夹中
			For i = 0 To UBound(aSavePathFileName)
				' 按路径文件名删除文件
				Call DoDelFile(aSavePathFileName(i))
			Next
   		 end if
	' 删除新闻
	Rs.delete
    Rs.update	
	Rs.movenext
	loop 
	Rs.Close
	set Rs=nothing	
	End If	
end if
conn.close
set conn=nothing  
response.redirect "usernews_ClassManage.asp"
%>



