<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/RemoteImg_save.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
call checkmanage("06")
Call OpenConn
session("newsuser")=request.form("newsuser")
session("come")=request.form("come")
dim keywords,img,firstimagename,i,scontent,content,newsuser,come,ok,city_oneid,city_twoid,city_threeid,city_one,city_two,city_three,c_oneid,c_twoid,c_threeid,c_one,c_two,c_three,pic_url,T_SaveImg,AspJpeg,T_UploadDir,m
%>
<%if request("subsave")="sz" Then%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1">
<tr> 
      <td height="80" colspan="3" align="center"> 
	  文章添加成功，请选择：<p>
        <a href="admin_article_add.asp">断续添加文章</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_article_upload.asp?photoid=<%=request("id")%>">立即批量上传本文章相册</a>
 </td>
    </tr></table>
<%response.end
End if
if request.form("subsave")="subok" Then

dim id,mystr
id=request.form("id")
mystr=trim(request.form("title")) '要处理的字符串 
title=Replace(mystr," ","") '过虑字符中间空格
keywords=Replace_Text(trim(request.form("keywords")))
img=request.form("img")

'得出地区名
city_oneid=strint(request.form("city_one"))
city_twoid=strint(request.form("city_two"))
city_threeid=strint(request.form("city_three"))
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

'得出栏目名
c_oneid=strint(request.form("c_oneid"))
c_twoid=strint(request.form("c_twoid"))
c_threeid=strint(request.form("c_threeid"))
if c_oneid>0 then
set rs=server.createobject("adodb.recordset")
rs.open "select name from DNJY_news_class where twoid=0 and threeid=0 and oneid="&c_oneid,conn,1,1
c_one=rs("name")
rs.close
if c_twoid>0 then
rs.open "select name from DNJY_news_class where oneid="&c_oneid&" and threeid=0 and twoid="&c_twoid,conn,1,1
c_two=rs("name")
rs.close
end if
if c_twoid>0 and c_threeid>0 then
rs.open "select name from DNJY_news_class where oneid="&c_oneid&" and twoid="&c_twoid&" and threeid="&c_threeid,conn,1,1
c_three=rs("name")
rs.close
end if
set rs=nothing
end If

if CheckIsEmpty(request("content"))=false then 
Call Alert ("请填写内容","-1")
else
content=trim(request("content"))
End If

newsuser=request.form("newsuser")
if newsuser="" then
newsuser="ip126"
end if
come=request.form("come")
if come="" then
come=web
end if
ok=request.form("ok")



set rs=server.createobject("adodb.recordset")
sql="select * from DNJY_news where (id is null)"
rs.open sql,conn,1,3
rs.addnew
if city_twoid="" then city_twoid=0
if city_threeid="" then city_threeid=0
rs("city_oneid")=city_oneid
rs("city_one")=city_one
rs("city_two")=city_two
rs("city_twoid")=city_twoid
rs("city_three")=city_three
rs("city_threeid")=city_threeid
rs("c_oneid")=c_oneid
rs("c_one")=c_one
rs("c_twoid")=c_twoid
rs("c_two")=c_two
rs("c_threeid")=c_threeid
rs("c_three")=c_three
rs("title")=title
rs("keywords")=keywords
rs("oStyle")=request.form("oStyle")
rs("oColor")=request.form("oColor")
rs("content")=content
rs("come")=come
rs("newsuser")=newsuser
rs("username")=request.cookies("administrator")
rs("newstop")=request.form("newstop")
rs("tuijian")=request("tuijian")
rs("ifshow")=request("ifshow")
rs("infotime")=now()
rs.update
Call CloseRs(rs)
set rs=server.createobject("adodb.recordset")
sql="select top 1 id from DNJY_news order by id"&DN_OrderType&""
rs.open sql,conn,1,1
Call Alert ("文章添加成功！","admin_article_addsave.asp?subsave=sz&id="&rs("id")&"")
Call CloseRs(rs)
Call CloseDB(conn)
response.end
Else
Call Alert ("抱歉，操作错误！","-1")
response.end
end if
%>