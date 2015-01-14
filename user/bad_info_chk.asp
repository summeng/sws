<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/tt_auto_lock.asp" -->
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
<HTML>
 <HEAD>
<title>不良信息举报</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="/<%=strInstallDir%>css/style.css">
 </HEAD>

 <BODY>
<%if CInt(trim(request.form("validate_codee")))<>CInt(Session("ttpSN")) then
response.write "验证码不对！"
Session("ttpSN")=""
response.end
end if
dim infoid,Bad_infolx,Bad_info,Bad_infomail,useradd
infoid=trim(request.Form("infoid"))
Bad_infolx=trim(request.Form("Bad_infolx"))
Bad_info=HTMLEncode(trim(request.Form("Bad_info")))
Bad_infomail=trim(request.Form("Bad_infomail"))
userip=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
city_oneid=strint(trim(request.Form("city_oneid")))
city_twoid=strint(trim(request.Form("city_twoid")))
city_threeid=strint(trim(request.Form("city_threeid")))
If userip="" Then 
userip=Request.ServerVariables("REMOTE_ADDR")
end If
Call OpenConn
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_Bad_info where ip='"&userip&"' and infoid="&cstr(infoid)&" order by addsj "&DN_OrderType&""
rs.open sql,conn,1,3
if rs.eof and rs.bof Then
useradd=True
Else
useradd=false
End If

If session("addcs")>5 Then
response.write "您连续举报次数太多，请以后再关注！<p><a href=""javascript:window.close()"">关闭窗口</a>"
response.end
end If
  
If useradd=False then
  if DateDiff("h",rs("addsj"),Now())<24 And useradd=false then
  response.write "您今天已举报过该信息，请不要重复提交！<p><a href=""javascript:window.close()"">关闭窗口</a>"
  response.end
  end if
 end if
Call CloseRs(rs)

set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_Bad_info"
rs.open sql,conn,1,3
rs.addnew
rs("infoid")=infoid
rs("Bad_infolx")=Bad_infolx
rs("Bad_info")=Bad_info
rs("Bad_infomail")=Bad_infomail
rs("addsj")=now()
rs("ip")=userip
rs("city_oneid")=city_oneid
rs("city_twoid")=city_twoid
rs("city_threeid")=city_threeid
rs.update
session("addcs")=session("addcs")+1
Call CloseRs(rs)
Call CloseDB(conn)
response.write "您的举报已记录，我们会及时审核处理！<p><a href=""javascript:window.close()"">关闭窗口</a>"
response.end
%>
 </BODY>
</HTML>
