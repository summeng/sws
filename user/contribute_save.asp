<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="../Include/RemoteImg_save.asp"-->
<!--#include file="../Include/err.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
  
<%
if lmkg2="1" then
call errdick(50)
response.end
end if
'========����ˮ=========
Dim AppealNum,AppealCount 
AppealNum=10 'ͬһIP60�����������ƴ���
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
response.write "<center>����̫Ƶ���ˣ�Ъһ����ɣ�</center>" 
response.end 
End If
'========����ˮEND=========

Call OpenConn
Dim vip
If request.cookies("dick")("username")<>"" Then vip=conn.Execute("Select vip From DNJY_user Where username='"&request.cookies("dick")("username")&"'")(0)
If contribute=0 Then
response.write "��վ������Ͷ����ڣ�"
Call CloseRs(rs)
Call CloseDB(conn)
response.end
ElseIf contribute>1 And request.cookies("dick")("username")="" Then
response.write "�οͲ���Ͷ�壡"
Call CloseRs(rs)
Call CloseDB(conn)
response.end
ElseIf contribute>2 And vip=0 Then
response.write "VIP��Ա����ȨͶ�壡"
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
mystr=trim(request.form("title")) 'Ҫ������ַ��� 
title=Replace(mystr," ","") '�����ַ��м�ո�
keywords=Replace_Text(trim(request.form("keywords")))
img=request.form("img")

'�ó�������
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

'�ó���Ŀ��
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
Call Alert ("����д����","-1")
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
If request.form("username")="" Then rs("username")="�ο�"
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
if Not rs.eof Then rs("jf")=rs("jf")+tgjf'���Ա���ӻ���
rs.update
Call CloseRs(rs)
end if
Call CloseDB(conn)
Call Alert ("�ύ�ɹ�����Ҫ������˲�����ʾ��","contribute.asp")
response.end
else
Call Alert ("��Ǹ����������","-1")
response.end
end if
%>