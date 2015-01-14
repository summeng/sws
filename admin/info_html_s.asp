<!--#include file="../conn.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
	margin-top: 16px;
}
-->
</style>
<%if asphtm=0 then
response.write "<script language=JavaScript>" & chr(13) & "alert('系统已关闭生成静态页功能，无法生成静态页！');" &"window.location='info_html.asp'" & "</script>"
response.end
end if  
dim username,leixin,biaoti,fbsj,userip,memo,hfcs,dianhua,mobile,userqq,email,xxmpname,dizhi,youbian,biaotixs,scjiage,jiage,sdmba,usernameid,namea,xxtp,CreateType
Dim str2,ServerURL,aa,objfso,htmout,http,k,wcolor,ncolor,cksj
id=trim(request("selectedid"))
Call OpenConn
Server.ScriptTimeOut=30000
CreateType=strint(Trim(Request("CreateType")))
If CreateType=1 Then
id=""
Dim TopNew
TopNew = strint(Trim(Request("TopNew")))
set rs=server.createobject("adodb.recordset")
Sql="Select top " & TopNew & " id From DNJY_info where Readinfo=0 and yz=1"
if jiaoyi=0 Then sql=sql&" and zt="&jiaoyi
if overdue=0 Then sql=sql&" and dqsj>"&DN_Now&""
sql=sql&" order by addsj "&DN_OrderType&""
rs.open sql,conn,1,1
do while not rs.eof
id=id+""&rs("id")&","
k=k+1
rs.movenext
if k>=TopNew then exit do
loop
Call CloseRs(rs)
id=StrReverse(Mid(StrReverse(id),2))
End if

if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('没有选择记录！');" &"window.location='info_html.asp'" & "</script>"
response.end
end if
str2=split(id,",")
for i=0 to ubound(str2)
set rs=server.createobject("adodb.recordset")
sql="select id,tj,fsohtm,city_oneid,city_twoid,city_threeid from [DNJY_info] where Readinfo=0 and yz=1 and id="&trim(str2(i))
rs.open sql,conn,1,3
if rs.eof Then
Response.Write ("<script language=javascript>alert('含有未审核或不存在ID号的信息！不能生成htm文件!');history.go(-1);</script>")
Call CloseRs(rs)
response.end
end if
rs("fsohtm")=1
rs.update
id=rs("id")
city_oneid=rs("city_oneid")
city_twoid=rs("city_twoid")
city_threeid=rs("city_threeid")

Response.Write "免费版无此功能！"
Response.end

Next
Call CloseRs(rs)
Call CloseDB(conn)
response.write "<meta http-equiv=refresh content=""1;URL=info_html.asp"">"
%>