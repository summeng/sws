<!--#include file="conn.asp"-->
<!--#include file=usercookies.asp-->
<!--#include file=top.asp-->
 <!--#include file="Include/err.asp"-->
<!--#include file="Include/RemoteImg_save.asp"-->
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
<link rel="stylesheet" type="text/css" href="css/css.css">
<title>添加信息</title>
</head>
<%if lmkg2="1" then
call errdick(50)
response.end
end if
Dim Classname,picture,Descrip,author,origin,jsid,Classid,Content,website,addtime,lhauthor,lhorigin,newsShared,pic_url,T_SaveImg,AspJpeg,T_UploadDir
session("author")=request.form("author")
session("origin")=request.form("origin")
username=request.cookies("dick")("username")
ClassName=Request.Form("ClassName")
title=request.form("txttitle")
picture=Request.Form("picture")
keywords=Request.Form("keywords")
Descrip=Request.Form("Descrip")
author=Request.Form("author")
origin=Request.Form("origin")
newsShared=Request.Form("newsShared")
Call OpenConn
set rs = server.CreateObject ("Adodb.recordset")
sql="select * From DNJY_userNewsClass where ClassName='"&ClassName&"'"
rs.open sql,conn,1,1
jsid=rs("id")
Classid=rs("ID")
Call CloseRs(rs)

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

If trim(request("d_content"))="" Then
Call Alert ("请填写内容","-1")
else
content=trim(request.Form("d_content"))
End If



'写入数据库
Set rs=CreateObject("Adodb.recordset")
sql="select * from DNJY_userNews"
rs.open sql,conn,1,3
rs.addnew
if city_twoid="" then city_twoid=0
if city_threeid="" then city_threeid=0
rs("city_oneid")=city_oneid
rs("city_one")=city_one
rs("city_twoid")=city_twoid
rs("city_two")=city_two
rs("city_threeid")=city_threeid
rs("city_three")=city_three
rs("username")=username
rs("title")=title
rs("ClassName")=ClassName
rs("ClassID")=ClassID
if Content<>"" then
rs("content")=content
end if
if keywords<>"" then
	rs("keywords")=keywords
else
	keywords=title&"_"&website
end if
if Descrip<>"" then
	rs("Descrip")=Descrip
else
	descrip=content
	descrip=CutStr(RemoveHTML(descrip),150,"")
end if

if author<>"" then
	rs("author")=author
else
	author=lhauthor
end if

if origin<>"" then
	rs("origin")=origin
else
	origin=lhorigin
end if
rs("addtime")=now()
rs("hits")=1
rs("newsShared")=newsShared
rs.update
id=rs("id")
Call CloseRs(rs)
Call CloseDB(conn)
Response.Write "<div align=center><br><br>保存成功！<p>是否继续添加此小类信息 <a href='UserNews_Add.asp?ClassName="&ClassName&"'>是</a> / <a href='UserNews_AddSelClass.asp'>否</a></div>"
%>
