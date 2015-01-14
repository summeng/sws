<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="config.asp"-->
<!--#include file=usercookies.asp-->
<!--#include file=Include/err.asp-->
<!--#include file=Include/ipt.asp-->
<!--#include file="Include/head.asp"-->
<!--#include file="Include/Function.asp"-->
<!--#include file="Include/label.asp"-->
<!--#include file=Include/md5.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
dim username,email,name,password,Regtime,idcard,Sex,dianhua,CompPhone,youbian,dizhi,http,Emap,weixin
dim maillist,rsg,sqlg,problem,answer1,answer2,answer_1,answer_2
username=request.cookies("dick")("username")
name=trim(request("name"))
email=trim(request("email"))
maillist=trim(request("maillist"))
idcard=trim(request("idcard"))
Sex=trim(request("Sex"))
dianhua=trim(request("dianhua"))
fax=trim(request.form("fax"))
CompPhone=trim(request("CompPhone"))
youbian=trim(request("youbian"))
qq=trim(request("qq"))
weixin=trim(request("weixin"))
dizhi=HTMLEncode(trim(request.form("dizhi")))
http=trim(request.form("http"))
city_oneid=strint(request.form("city_one"))
city_twoid=strint(request.form("city_two"))
city_threeid=strint(request.form("city_three"))
type_oneid=strint(request.form("type_one"))
type_twoid=strint(request.form("type_two"))
type_threeid=strint(request.form("type_three"))
Emap=request.form("Emap")
problem=trim(request("problem"))
answer1=trim(request("answer1"))
answer2=trim(request("answer2"))
answer_1=md5(answer1)
answer_2=md5(answer2)
Call OpenConn
	If city_oneid>0 Then city_one=Conn.Execute("select city from DNJY_city where twoid=0 and id="&city_oneid)(0)
	If city_twoid>0 Then city_two=Conn.Execute("select city from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid)(0)
	If city_threeid<>0 Then city_three=Conn.Execute("select city from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid)(0)
	
	If type_oneid>0 Then type_one=Conn.Execute("select name from DNJY_type where twoid=0 and id="&type_oneid)(0)
	If type_twoid>0 Then type_two=Conn.Execute("select name from DNJY_type where id="&type_oneid&" and threeid=0 and twoid="&type_twoid)(0)
	If type_threeid<>0 Then type_three=Conn.Execute("select name from DNJY_type where id="&type_oneid&" and twoid="&type_twoid&" and threeid="&type_threeid)(0)

    set rsg=server.createobject("adodb.recordset")
	sqlg="select * from [DNJY_user] where email<>'' and email='"&email&"' and username<>'"&username&"'"
	rsg.open sqlg,conn,1,1
	if not rsg.eof then	
	call errdick(23)
    Call CloseRs(rsg)
    Call CloseDB(conn)
    response.end
	end if
	set rsg=nothing	

if email="" and maillist=1 then
call errdick(34)
response.end
end if
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,3
rs("city_oneid")=city_oneid
rs("city_twoid")=city_twoid
rs("city_threeid")=city_threeid
rs("city_one")=city_one
rs("city_two")=city_two
rs("city_three")=city_three

rs("type_oneid")=type_oneid
rs("type_twoid")=type_twoid
rs("type_threeid")=type_threeid
rs("type_one")=type_one
rs("type_two")=type_two
rs("type_three")=type_three

rs("name")=name
rs("email")=email
rs("maillist")=maillist
rs("idcard")=idcard
rs("Sex")=Sex
rs("dianhua")=dianhua
rs("CompPhone")=CompPhone
if youbian<>"" then
rs("youbian")=youbian
end if
rs("dizhi")=dizhi
rs("fax")=fax
rs("http")=http
rs("qq")=qq
rs("weixin")=weixin
rs("Emap")=Emap
if problem<>"" then
rs("problem")=problem
rs("answer")=answer_1
end if
rs.update
Call CloseRs(rs)
Call CloseDB(conn)
response.write "<script language=JavaScript>" & chr(13) & "alert('操作成功！');" &"window.location='user_data.asp?"&C_ID&"'" & "</script>"
response.end
%>
