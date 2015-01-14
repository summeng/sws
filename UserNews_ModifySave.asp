<!--#include file="conn.asp"-->
<!--#include file=top.asp-->
<!--#include file="Include/err.asp"-->
<!--#include file=usercookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<title>修改信息</title>
</head>
<%Dim isse,Classname,se_Classname,page,searchkeys,selecttype,addtime,Descrip,author,origin,ClassID,Content,website,lhauthor,lhorigin,newsShared,T_SaveImg,AspJpeg,T_UploadDir
isse=request.QueryString("isse")
Classname=Request.Form("Classname")
se_Classname=Request.QueryString("se_Classname")
page=Request.QueryString("page")
ID=Request.QueryString("id")
searchkeys=Request.QueryString("searchkeys")
selecttype=request.QueryString("selecttype")
title=request.Form("txttitle")
addtime=Request.Form("addtime")
keywords = Request.Form("keywords")
Descrip=Request.Form("Descrip")
author=Request.Form("author")
origin=Request.Form("origin")
newsShared=Request.Form("newsShared")

Call OpenConn
city_oneid=strint(request.form("city_oneid"))
city_twoid=strint(request.form("city_twoid"))
city_threeid=strint(request.form("city_threeid"))
if city_oneid>0 then
set rs=server.createobject("adodb.recordset")
rs.open "select city from DNJY_city where twoid=0 and threeid=0 and id="&city_oneid,conn,1,1
city_one=rs("city")
rs.close
if city_twoid>0 then
rs.open "select city from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid,conn,1,1
city_two=rs("city")
rs.close
end if
if city_twoid>0 and city_threeid>0 then
rs.open "select city from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid,conn,1,1
city_three=rs("city")
rs.close
end if
set rs=nothing
end If

'得类别id
set rs = server.CreateObject ("Adodb.recordset")
sql="select * From DNJY_userNewsClass where Classname='"& Classname &"'"
rs.open sql,conn,1,1
ClassID=rs("ID")
Call CloseRs(rs)

If trim(request("d_content"))="" Then
Call Alert ("请填写内容","-1")
else
content=trim(request.Form("d_content"))
End If



'修改信息
Set rs=CreateObject("Adodb.recordset")
sql="select * from DNJY_userNews where id="&id
rs.open sql,conn,1,3
if city_twoid="" then city_twoid=0
if city_threeid="" then city_threeid=0
rs("city_oneid")=city_oneid
rs("city_one")=city_one
rs("city_twoid")=city_twoid
rs("city_two")=city_two
rs("city_threeid")=city_threeid
rs("city_three")=city_three
rs("title")=title
rs("Classname")=Classname
rs("ClassID")=ClassID
if Content<>"" then
	rs("content")=content
end if
rs("keywords") = keywords
if keywords="" then
	keywords=title&"_"&website
end if

rs("author") = author
if author="" then
	author=lhauthor
end if

rs("origin") = origin
if origin="" then
	origin=lhorigin
end if

rs("Descrip")=Descrip
if Descrip="" then
	descrip=content
	descrip=CutStr(RemoveHTML(descrip),150,"")
end if

rs("addtime")=addtime
rs("newsShared")=newsShared
rs.update
Call CloseRs(rs)
Call CloseDB(conn)
if isse=0 then
Response.Write "<div align=center><br><br>修改成功！<p>是否继续修改此小类信息 <a href='UserNews_Modifylist.asp?isse="&isse&"&Classname="&Classname&"&page="&page&"'>是</a> / <a href='UserNews_ModSelClass.asp'>否</a></div>"
end if
if isse=1 then
	if selecttype=2 then
			response.write"<SCRIPT language=JavaScript>alert('修改成功！');"
			response.write"location.href ='UserNews_Search.asp'</SCRIPT>"
	end if
	if selecttype=1 then
	Response.Write "<div align=center><br><br>修改成功！<p>是否继续修改此查询信息 <a   href='UserNews_Modifylist.asp?isse="&isse&"&searchkeys="&searchkeys&"&selecttype="&selecttype&"&se_ClassName="&se_ClassName&"&page="&page&"'>是</a> / <a href='UserNews_ModSelClass.asp'>否</a></div>"
	end if
end if
%>
