<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file=cookies.asp-->
<!--#include file=../Include/err.asp-->
<!--#include file=../Include/ipt.asp-->
<!--#include file=../Include/mail.asp-->
<!--#include file=../Include/md5.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Call OpenConn
dim id,username,email,name,password,idcard,rs2,sql2,Sex,dianhua,fax,CompPhone,qq,weixin,mpname,youbian,dizhi,maillist,city_oneid,city_twoid,city_threeid,city_one,city_two,city_three,type_oneid,type_twoid,type_threeid,type_one,type_two,type_three,useryz,mailyz
username=trim(request("username"))
name=trim(request("name"))
useryz=trim(request("useryz"))
mailyz=trim(request("mailyz"))
If trim(request("password"))<>"" then
password=md5(trim(request("password")))
End if
email=trim(request("email"))
maillist=trim(request("maillist"))
idcard=trim(request("idcard"))
Sex=trim(request("Sex"))
dianhua=trim(request("dianhua"))
fax=trim(request("fax"))
mpname=trim(request("mpname"))
youbian=trim(request("youbian"))
dizhi=trim(request("dizhi"))
CompPhone=trim(request("CompPhone"))
qq=trim(request("qq"))
weixin=trim(request("weixin"))
city_oneid=strint(request.form("city_one"))
city_twoid=strint(request.form("city_two"))
city_threeid=strint(request.form("city_three"))
type_oneid=strint(request.form("type_one"))
type_twoid=strint(request.form("type_two"))
type_threeid=strint(request.form("type_three"))
    If city_oneid>0 Then city_one=Conn.Execute("select city from DNJY_city where twoid=0 and id="&city_oneid)(0)
	If city_twoid>0 Then city_two=Conn.Execute("select city from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid)(0)
	If city_threeid<>0 Then city_three=Conn.Execute("select city from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid)(0)
	If type_oneid>0 Then type_one=Conn.Execute("select name from DNJY_type where twoid=0 and id="&type_oneid)(0)
	If type_twoid>0 Then type_two=Conn.Execute("select name from DNJY_type where id="&type_oneid&" and threeid=0 and twoid="&type_twoid)(0)
	If type_threeid<>0 Then type_three=Conn.Execute("select name from DNJY_type where id="&type_oneid&" and twoid="&type_twoid&" and threeid="&type_threeid)(0)	
chkdick(name)


set rs2=server.createobject("adodb.recordset")
sql2="select email from [DNJY_user]  where email<>'' and email='"&email&"' and username<>'"&username&"'"
rs2.open sql2,conn,1,3
if not rs2.eof then
Response.Write ("<script language=javascript>alert('此信箱已存在，请重新输入!');history.go(-1);</script>")
 response.end
 end If 
Call CloseRs(rs2)

set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_user] where username='"&trim(request("username"))&"'"
rs.open sql,conn,1,3
rs("name")=name
If password<>"" Then rs("password")=password
rs("email")=email
rs("maillist")=maillist
rs("idcard")=idcard
rs("Sex")=Sex
rs("dianhua")=dianhua
rs("fax")=fax
rs("CompPhone")=CompPhone
rs("qq")=qq
rs("weixin")=weixin
rs("mpname")=mpname
if youbian<>"" then
rs("youbian")=youbian
end if
rs("dizhi")=dizhi
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
rs("useryz")=useryz
'If useryz=1 Then rs("mailyz")=1
rs.update
Call CloseRs(rs)

Call CloseDB(conn)

response.Write "<script language=javascript>alert('修改成功!');location.replace('user_editdata.asp?username="&username&"');</script>"
'call cl()
sub cl()%>
<body onLoad="setTimeout(window.close, 2000)">
<%end sub%>
