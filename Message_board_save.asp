<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="Include/err.asp"-->
<!--#include file="config.asp"-->
<!--#include file="Include/head.asp"-->
<!--#include file="Include/Function.asp"-->
<!--#include file="Include/label.asp"-->
<!--#include file="Include/tt_auto_lock.asp" -->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%if not ChkPost() then 
response.write "<center>禁止站外提交或访问"
response.end 
end if
dim reg1,ses,ttdv
reg1=request.cookies("dick")("reg1")
ses=DateDiff("s",reg1,now())
	city_oneid=Strint(Request.Form("city_one"))
	city_twoid=Strint(Request.Form("city_two"))
	city_threeid=Strint(Request.Form("city_three"))
if city_oneid<=0 then
response.write "<script language=JavaScript>" & chr(13) & "alert('请选择您所在地区！');" & "history.back()" & "</script>"
response.end
end if

if ses<60 then
response.write "<script language=JavaScript>" & chr(13) & "alert('你的留言间隔是60秒，请等会再次发布！');" & "history.back()" & "</script>"
response.end
end If

if codekg1=1 then
if Trim(Request.Form("wenti"))=Empty Or trim(request.form("daan"))<>trim(request.form("wenti")) then
call errdick(44)
response.end
end if
end If

if codekg2=1 then
if trim(request.form("yzm"))<>Session("pSN") then
call errdick(39)
Session("pSN")=""
response.end
end if
end if

if codekg3=1 then
if lcase(Request.Form("code"))<>lcase(Session("GetCode")) then
call errdick(40)
Session("GetCode")=""
response.end
end If
end If

if codekg4=1 then
ttdv=cint(trim(request("ttdv")))+1
if ttdv<>cint(datepart("w",date())) Then
call errdick(41)
response.end
end If
end If

if codekg5=1 then
if CInt(trim(request.form("validate_codee")))<>CInt(Session("ttpSN")) then
call errdick(45)
Session("ttpSN")=""
response.end
end if
end If


if trim(request.form("memo"))="" then
call errdick(4)
response.end
end If

Call OpenConn
	city_one=Conn.Execute("select city from DNJY_city where twoid=0 and id="&city_oneid)(0)
	If city_twoid>0 Then city_two=Conn.Execute("select city from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid)(0)
	If city_threeid<>0 Then city_three=Conn.Execute("select city from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid)(0)

set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_gbook order by fbsj "&DN_OrderType&""
rs.open sql,conn,1,3
rs.addnew
if city_oneid="" then city_oneid=0
if city_twoid="" then city_twoid=0
if city_threeid="" then city_threeid=0
rs("city_oneid")=city_oneid
rs("city_twoid")=city_twoid
rs("city_threeid")=city_threeid
rs("city_one")=city_one
rs("city_two")=city_two
rs("city_three")=city_three
rs("username")=request.cookies("dick")("username")
rs("lx")=trim(request.Form("gbook"))
rs("gbook1")=HTMLEncode(trim(request.Form("memo")))
If request.cookies("dick")("username")<>"" then
rs("known")=trim(request.Form("known"))
End if
rs("fbsj")=now()
rs.update
Response.cookies("dick")("reg1")=now()
Call CloseRs(rs)
Call CloseDB(conn)
session("book_error")=session("book_error")+1  '一个会话期内限制留言
response.write "<script language=JavaScript>" & chr(13) & "alert('留言成功，需要管理员回复后才可以显示在网站上(对会员本人无限制)！');" &"window.location='"&strInstallDir&"index.asp?"&CT_ID&"'" & "</script>"
response.write "<meta http-equiv=refresh content=""3;URL="&strInstallDir&"index.asp?ts=1&"&CT_ID&""">"
%>
