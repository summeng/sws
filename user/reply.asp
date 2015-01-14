<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="../Include/mail.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%if not ChkPost() then 
response.write "禁止站外提交或访问"
response.end 
end if
dim username,nm,userip,biaoti,rsmes,email
biaoti=trim(request("biaoti"))
id=trim(request("id"))
nm=Request.form("nm")
username=request("username")
email=request("email")
city_oneid=request("city_oneid")
city_twoid=request("city_twoid")
city_threeid=request("city_threeid")
if Request.form("r1")="0" And trim(request("xxusername"))="游客" Then
response.write "<li>该用户未注册，不能直接回复，请选择'发到发布人的信箱'！<a href='javascript:history.back(-1);'>返回</a>"
response.end
end If

if Request.form("r1")="1" And (request.cookies("dick")("username")="" or request.cookies("dick")("domain")="" or request.cookies("dick")("id")="")   Then
response.write "<br>"
response.write "<li>发邮件方式需要先登录，您还没有登陆！"
response.write "<meta http-equiv=refresh content=""2;URL=../login.asp?"&C_ID&""">"
response.end
end If
if Request.form("validate_codee")=""  Then
response.write "<li>验证码没有填写！<a href='javascript:history.back(-1);'>返回</a>"
cl
response.end
end If

if not isnumeric(id) or id="" then
response.write "<li>参数错误！"
cl
response.end
end if
if nm="on" Then
username="匿名"
call dick()
elseif (request.cookies("dick")("username")="" or request.cookies("dick")("domain")="" or request.cookies("dick")("id")="") And newspl=1 then 
response.write "<br>"
response.write "<li>您还没有登陆！"
response.write "<meta http-equiv=refresh content=""2;URL=../login.asp?"&C_ID&""">"
response.End
elseif (request.cookies("dick")("username")="" or request.cookies("dick")("domain")="" or request.cookies("dick")("id")="") And newspl=2 then 
Response.Write ("<script language=javascript>alert('禁止新评论!');history.go(-1);</script>")
response.end
    else 
    if request("dick")="chk" then
    call dick()
    response.end
    end if

end if
%>

<%
sub dick()%>

<%dim neirong
if len(trim(request.form("neirong")))<1 then
response.write "<li>回复内容没有填写！"
cl
response.end
end If

if len(trim(request.form("neirong")))>200 then
response.write "<li>回复内容不能大于200个字符！"
cl
response.end
end If

if CInt(trim(request.form("validate_codee")))<>CInt(Session("ttpSN")) then
response.write "验证码不对！"
Session("ttpSN")=""
cl
response.end
end if

Call OpenConn
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_reply "
rs.open sql,conn,1,3
rs.addnew
rs("username")=request.cookies("dick")("username")
rs("neirong")=trim(request("neirong"))
rs("xxusername")=trim(request("xxusername"))
rs("xxid")=id
userip=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If userip="" Then 
userip=Request.ServerVariables("REMOTE_ADDR")
end if
rs("ip")=userip
rs("city_oneid")=city_oneid
rs("city_twoid")=city_twoid
rs("city_threeid")=city_threeid
rs("hfsj")=now()
If plsh=0 Then rs("hfyz")=1
rs.update
Call CloseRs(rs)

'向信息发布者发送短信
	sql="select * from DNJY_Message"
	set rsmes=server.createobject("adodb.recordset")
	rsmes.open sql,conn,1,3
    rsmes.AddNew
  	rsmes("name")=trim(request("xxusername"))
	neirong="有人对您发布的信息【"&biaoti&"】很感兴趣，并做了回复，请及时查看！如果查看不到回复，可能是管理员未审核或审核未通过。"
	rsmes("neirong")=neirong
  	rsmes("riqi")=now()
 	rsmes("fname")=username
rsmes.Update
rsmes.close
set rsmes=Nothing

'向信息发布者发送邮件
dim mailbody,Jmail
mailbody="有人对您在【<a href='http://"&web&"'>"&webname&"</a>】发布的信息【"&biaoti&"】很感兴趣，并做了回复，请及时查看！如果查看不到回复，可能是管理员未审核或审核未通过。可联系回复人邮箱："&request("mymail")&"。【本邮件由系统自动发送，请勿直接回复】"
'邮件发送
Dim HostName,E_Subject,Information,S_Type,C_M_Chk,e_info
Email=email
HostName=webname
E_Subject="有人对您在"&webname&"发布的信息【"&biaoti&"】很感兴趣"
Information=mailbody
S_Type=True
C_M_Chk=True
If ChkMail(Email)=true Then Call DNJYEmail(Email,HostName,E_Subject,Information,S_Type,C_M_Chk)
'邮件发送END

Set Jmail=Nothing



Conn.Execute("Update DNJY_info Set hfcs=hfcs+1 where id="&cstr(id))
Conn.Execute("Update DNJY_info Set hfsj="&DN_Now&" where id="&cstr(id))
If plsh=0 Then
response.write "<li>回复成功！"
else
response.write "<li>回复成功！管理员审核后可见。"
End if
cl
end sub
%>
<%sub cl()
response.write "<meta http-equiv=refresh content=""2;URL=../"&x_path("",id,"",Readinfo)&""">"
end Sub
Call CloseDB(conn)%>