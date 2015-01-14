<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="../Include/RemoteImg_save.asp"-->
<!--#include file="../Include/err.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
  
<%
if lmkg2="1" then
call errdick(50)
response.end
end if
'========防灌水=========
Dim AppealNum,AppealCount 
AppealNum=10 '同一IP60秒内请求限制次数
AppealCount=Request.Cookies("AppealCount") 
If AppealCount="" Then 
response.Cookies("AppealCount")=1 
AppealCount=1 
response.cookies("AppealCount").expires=dateadd("s",60,now()) 
Else 
response.Cookies("AppealCount")=AppealCount+1 
response.cookies("AppealCount").expires=dateadd("s",60,now()) 
End If 
if int(AppealCount)>int(AppealNum) then 
response.write "<center>访问太频繁了，歇一会儿吧！</center>" 
response.end 
End If
'========防灌水END=========

Call OpenConn
Dim vip
If request.cookies("dick")("username")<>"" Then vip=conn.Execute("Select vip From DNJY_user Where username='"&request.cookies("dick")("username")&"'")(0)
If contribute=0 Then
response.write "本站不开放投稿入口！"
Call CloseRs(rs)
Call CloseDB(conn)
response.end
ElseIf contribute>1 And request.cookies("dick")("username")="" Then
response.write "游客不能投稿！"
Call CloseRs(rs)
Call CloseDB(conn)
response.end
ElseIf contribute>2 And vip=0 Then
response.write "VIP会员才有权投稿！"
Call CloseRs(rs)
Call CloseDB(conn)
response.end
End if
session("newsuser")=request.form("newsuser")
session("come")=request.form("come")
dim keywords,img,i,scontent,content,newsuser,come,ok,rsbigc,city_oneid,city_twoid,city_threeid,city_one,city_two,city_three,type_oneid,type_twoid,type_threeid,type_one,type_two,type_three,c_oneid,c_twoid,c_threeid,c_one,c_two,c_three,pic_url,T_SaveImg,AspJpeg,T_UploadDir
function txttourl(tekst)
tekst_temp = tekst
tekst_temp = replace(tekst_temp,"<","&lt;")
tekst_temp = replace(tekst_temp,">","&gt;")
tekst_temp = replace(tekst_temp,"chr(13)","<br>")
tekst_temp = replace(tekst_temp,"#","<br>")
tekst_temp = replace(tekst_temp,chr(13),"<br>")
tekst_temp = replace(tekst_temp,"[","<b>")
tekst_temp = replace(tekst_temp,"]","</b>")
txttourl = tekst_temp
end Function

if request.form("subsave")="subok" Then

dim id,SmallClassID,mystr
id=request.form("id")
SmallClassID=request.form("SmallClassID")
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

If trim(request("content"))="" Then
Call Alert ("请填写内容","-1")
else
content=trim(request("content"))
End If

newsuser=request.form("newsuser")
if newsuser="" then
newsuser="ip126"
end if
come=request.form("come")
ok=request.form("ok")



set rs=server.createobject("adodb.recordset")
sql="select * from DNJY_News where (id is null)"
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
rs("c_two")=c_two
rs("c_twoid")=c_twoid
rs("c_three")=c_three
rs("c_threeid")=c_threeid
rs("title")=title
rs("keywords")=keywords
rs("oStyle")=request.form("oStyle")
rs("oColor")=request.form("oColor")
rs("content")=content
rs("come")=come
rs("newsuser")=newsuser
rs("username")=request.form("username")
If request.form("username")="" Then rs("username")="游客"
rs("newstop")=request.form("newstop")
rs("tuijian")=request("tuijian")
rs("ifshow")=0
rs("infotime")=now()
rs.update
Call CloseRs(rs)
If request.cookies("dick")("username")<>"" Then
set rs=server.createobject("adodb.recordset")
sql="select * from DNJY_user where username='"&request.cookies("dick")("username")&"'"
rs.open sql,conn,1,3
if Not rs.eof Then rs("jf")=rs("jf")+tgjf'向会员增加积分
rs.update
Call CloseRs(rs)
end if
Call CloseDB(conn)
Call Alert ("提交成功！需要网管审核才能显示。","contribute.asp")
response.end
else
Call Alert ("抱歉，操作错误！","-1")
response.end
end if
%>