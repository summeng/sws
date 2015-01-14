<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=CONfig.ASP-->
<!--#include file=usercookies.asp-->
<!--#include file=Include/err.asp-->
<!--#include file=Include/ipt.asp-->
<!--#include file="Include/Function.asp"-->
<!--#include file="Include/label.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
dim username,sdname,sdjj,sdfg,sdkg,mpname,mpjj,mpfg,mpkg,name,email,dianhua,CompPhone,youbian,fax,qq,dizhi,http,zhuying,city_oneid,city_twoid,city_threeid,type_oneid,type_twoid,type_threeid,city_one,city_two,city_three,type_one,type_two,type_three
username=request.cookies("dick")("username")
sdname=HTMLEncode(trim(request.form("sdname")))
sdjj=HTMLEncode(trim(request.form("sdjj")))
sdfg=trim(request.form("sdfg"))
sdkg=trim(request.form("sdkg"))
mpname=HTMLEncode(trim(request.form("mpname")))
mpjj=HTMLEncode(trim(request.form("mpjj")))
mpfg=trim(request.form("mpfg"))
mpkg=trim(request.form("mpkg"))
name=HTMLEncode(trim(request.form("name")))
email=trim(request.form("email"))
dianhua=trim(request.form("dianhua"))
CompPhone=trim(request.form("CompPhone"))
fax=trim(request.form("fax"))
qq=trim(request.form("qq"))
youbian=trim(request.form("youbian"))
If trim(request.form("youbian"))="" Then youbian=0 
dizhi=HTMLEncode(trim(request.form("dizhi")))
http=trim(request.form("http"))
zhuying=HTMLEncode(trim(request.form("zhuying")))
city_oneid=strint(request.form("city_one"))
city_twoid=strint(request.form("city_two"))
city_threeid=strint(request.form("city_three"))
type_oneid=Strint(Request.Form("type_one"))
type_twoid=Strint(Request.Form("type_two"))
type_threeid=Strint(Request.Form("type_three"))
Call OpenConn
set rs=server.createobject("adodb.recordset")
sql="select sdname from [DNJY_user] where username<>'"&username&"' and sdname='"&sdname&"'"
rs.open sql,conn,1,1
If not rs.eof Then
response.write "<script LANGUAGE='javascript'>alert('温馨小提示：您选择的网店名称太热门了，已有人入驻，请另选!');history.go(-1);</script>"
Call CloseRs(rs)
Call CloseDB(conn)
Response.End
End If

	city_one=Conn.Execute("select city from DNJY_city where twoid=0 and id="&city_oneid)(0)
	If city_twoid>0 Then city_two=Conn.Execute("select city from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid)(0)
	If city_threeid<>0 Then city_three=Conn.Execute("select city from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid)(0)
	type_one=Conn.Execute("select name from DNJY_hy_type where twoid=0 and id="&type_oneid)(0)
	If type_twoid<>0 Then type_two=Conn.Execute("select name from DNJY_hy_type where id="&type_oneid&" and threeid=0 and twoid="&type_twoid)(0)
	If type_threeid<>0 Then type_three=Conn.Execute("select name from DNJY_hy_type where id="&type_oneid&" and twoid="&type_twoid&" and threeid="&type_threeid)(0)

set rs=server.createobject("adodb.recordset")
sql="select sdname,sdjj,sdfg,sdkg,mpname,mpjj,mpfg,mpkg,name,email,dianhua,CompPhone,fax,qq,youbian,dizhi,http,zhuying,city_oneid,city_twoid,city_threeid,type_oneid,type_twoid,type_threeid,city_one,city_two,city_three,type_one,type_two,type_three,notification from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,3
rs("sdname")=sdname
rs("sdjj")=sdjj
rs("sdfg")=sdfg
rs("sdkg")=sdkg
rs("mpname")=mpname
rs("mpjj")=mpjj
rs("mpfg")=mpfg
rs("mpkg")=mpkg
rs("name")=name
rs("email")=email
rs("dianhua")=dianhua
rs("CompPhone")=CompPhone
rs("fax")=fax
rs("qq")=qq
rs("youbian")=youbian
rs("dizhi")=dizhi
rs("http")=http
rs("zhuying")=zhuying
if city_twoid="" then city_twoid=0
if city_threeid="" then city_threeid=0
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
rs("notification")="等待管理员审核..."
rs.update
Call CloseRs(rs)
Call CloseDB(conn)
response.write "<script language=JavaScript>" & chr(13) & "alert('操作成功！申请(修改)店企需经重新审核才能显示。');" &"window.location='usersdeditzl.asp'" & "</script>"
response.end
%>
