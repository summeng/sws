<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file=../Include/Function.asp-->
<!--#include file=cookies.asp-->
<!--#include file="../Include/RemoteImg_save.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("04")%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" href="../css/css.css">
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
.style1 {color: #FFFFFF;
	font-weight: bold;
}
.style2 {color: #FFFFFF}
-->
</style>
<title>修改信息</title>
</head>
<%Dim isse,Classname,se_Classname,page,searchkeys,selecttype,addtime,Descrip,author,origin,ClassID,Content,website,lhauthor,lhorigin,ifshow,newstop,tj,T_SaveImg,AspJpeg,T_UploadDir
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
tj=Request.Form("tj")
newstop=Request.Form("newstop")
ifshow=Request.Form("ifshow")
'得类别id
set rs = server.CreateObject ("Adodb.recordset")
sql="select * From DNJY_userNewsClass where Classname='"& Classname &"'"
rs.open sql,conn,1,1
ClassID=rs("ID")
Call CloseRs(rs)

city_oneid=HtmlEncode(request.form("city_one"))
city_twoid=HtmlEncode(request.form("city_two"))
city_threeid=HtmlEncode(request.form("city_three"))
if city_oneid<>"" then
set rs=server.createobject("adodb.recordset")
rs.open "select city from DNJY_city where twoid=0 and threeid=0 and id="&city_oneid,conn,1,1
city_one=rs("city")
rs.close
if city_twoid<>"" then
rs.open "select city from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid,conn,1,1
city_two=rs("city")
rs.close
end if
if city_twoid<>"" and city_threeid<>"" then
rs.open "select city from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid,conn,1,1
city_three=rs("city")
rs.close
end If
set rs=nothing
end If

If trim(request("d_content"))="" Then
Call Alert ("请填写内容","-1")
else
content=trim(request.Form("d_content"))
End If



'修改信息
Set rs=CreateObject("Adodb.recordset")
sql="select * from DNJY_userNews where id="&id
rs.open sql,conn,1,3
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
rs("tj")=tj
rs("newstop")=newstop
rs("ifshow")=ifshow
if city_oneid="" then city_oneid=0
rs("city_oneid")=city_oneid
rs("city_one")=city_one
if city_twoid="" then city_twoid=0
rs("city_twoid")=city_twoid
rs("city_two")=city_two
if city_threeid="" then city_threeid=0
rs("city_threeid")=city_threeid
rs("city_three")=city_three
rs.update
rs.close
set rs = nothing


Call CloseDB(conn)
Call Alert ("修改成功!","UserNews_Modifylist.asp?isse=0&ClassName="&ClassName&"")
%>
