<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file=cookies.asp-->
<!--#include file=../Include/err.asp-->
<!--#include file=../Include/ipt.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<%Call OpenConn
dim username,sdname,sdjj,sdfg,sdkg,mpname,mpjj,mpfg,mpkg,id,http,zhuying,type_oneid,type_twoid,type_threeid,type_one,type_two,type_three,sdgd,mpgd
id=trim(request("id"))
if not isnumeric(id) or id="" then
response.write "<li>参数错误！"
response.end
end If
username=trim(request("username"))
sdname=trim(request("sdname"))
sdjj=trim(request("sdjj"))
sdkg=trim(request("sdkg"))
sdfg=trim(request("sdfg"))
sdgd=trim(request("sdgd"))
mpname=trim(request("mpname"))
mpjj=trim(request("mpjj"))
mpkg=trim(request("mpkg"))
mpfg=trim(request("mpfg"))
mpgd=trim(request("mpgd"))
http=trim(request("http"))
zhuying=trim(request("zhuying"))
type_oneid=Strint(Request.Form("type_one"))
type_twoid=Strint(Request.Form("type_two"))
type_threeid=Strint(Request.Form("type_three"))
	type_one=Conn.Execute("select name from DNJY_hy_type where twoid=0 and id="&type_oneid)(0)
	If type_twoid<>0 Then type_two=Conn.Execute("select name from DNJY_hy_type where id="&type_oneid&" and threeid=0 and twoid="&type_twoid)(0)
	If type_threeid<>0 Then type_three=Conn.Execute("select name from DNJY_hy_type where id="&type_oneid&" and twoid="&type_twoid&" and threeid="&type_threeid)(0)

If sdname<>"" then
set rs=server.createobject("adodb.recordset")
sql="select sdname from [DNJY_user] where username<>'"&username&"' and sdname='"&sdname&"'"
rs.open sql,conn,1,1
If not rs.eof Then
response.write "<script LANGUAGE='javascript'>alert('温馨小提示：您选择的网店名称太热门了，已有人入驻，请另选!');history.go(-1);</script>"
Call CloseRs(rs)
Call CloseDB(conn)
Response.End
End If
End If

set rs=server.createobject("adodb.recordset")
sql="select username,sdname,sdjj,sdfg,sdkg,mpname,mpjj,mpfg,mpkg,dpdata,http,zhuying,type_oneid,type_twoid,type_threeid,type_one,type_two,type_three,sdgd,mpgd,notification from [DNJY_user] where id="&id
rs.open sql,conn,1,3
rs("sdname")=sdname
rs("sdjj")=sdjj
rs("sdkg")=sdkg
rs("sdfg")=sdfg
rs("sdgd")=sdgd
rs("mpname")=mpname
rs("mpjj")=mpjj
rs("mpkg")=mpkg
rs("mpfg")=mpfg
rs("mpgd")=mpgd
rs("http")=http
rs("zhuying")=zhuying
rs("type_oneid")=type_oneid
rs("type_twoid")=type_twoid
rs("type_threeid")=type_threeid
rs("type_one")=type_one
rs("type_two")=type_two
rs("type_three")=type_three
if request("gxsj")=1 then
rs("dpdata")=now()
end If
rs("notification")=trim(request("notification"))
rs.update
Call CloseRs(rs)
Call CloseDB(conn)
response.write "审核成功！！"
call cl()
%>
<%sub cl()%>
<body onLoad="setTimeout(window.close, 2000)">
<%end sub%>