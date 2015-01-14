<!--#include file="conn.asp"-->
<!--#include file=top.asp-->
<!--#include file=Include/err.asp-->
<!--#include file=source.asp-->
<!--#include file=Include/md5.asp-->
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
if Instr(Request.ServerVariables("HTTP_REFERER"),""&addinfo&"")=0 then
response.redirect ""&addinfo&"?"&C_ID&"" 
end If
if addxinxi=0 then
response.redirect "login.asp"
end If
if lmkg2="1" then
call errdick(50)
response.end
end if
if addip<>"0" then
if ip<>"" then
addip=split(cstr(ip),"@")
for N=0 to UBound(addip)
if instr(Request.serverVariables("REMOTE_ADDR"),addip(n))>0 then
errdick(43)
response.end
end if
next
end if
end If
%>

<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css.css">
<style type="text/css">
<!--
body {
	background-image: url();
}
-->
</style></head>
<body topmargin="0" leftmargin="0">
<table width="980"  border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td width="214" valign="top"><div align="center">
       
    </div></td>
    <td width="760" valign="top"><table width="760" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
      <tr>
        <td height="4" bgcolor="#e6e6e6"></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF"><table width="760" height="329" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#111111" bgcolor="#FFFFFF" style="border-collapse: collapse">
          <tr>
            <td width="760" valign="top" height="329">　
<%dim biaoti,scjiage,jiage,wcolor,ncolor,memo,a,b,d,name,dianhua,CompPhone,weixin,email,xxmpname,dizhi,youbian,sdays,yz,arrkillword,arrkillmemo,arrkillname,arrkilldizhi,j,Enable,ttdv,delpas,mpname,cksj

biaoti=HTMLEncode(trim(request.Form("biaoti")))
keywords=HTMLEncode(Replace_Text(trim(request.Form("keywords"))))
name=HTMLEncode(trim(request.Form("name")))
city_oneid=Strint(Request.Form("city_one"))
city_twoid=Strint(Request.Form("city_two"))
city_threeid=Strint(Request.Form("city_three"))
type_oneid=Strint(Request.Form("type_one"))
type_twoid=Strint(Request.Form("type_two"))
type_threeid=Strint(Request.Form("type_three"))
scjiage=trim(request.Form("scjiage"))
jiage=trim(request.Form("jiage"))
wcolor=trim(request.Form("wcolor"))
ncolor=trim(request.Form("ncolor"))
If trim(request.Form("memo"))="" Then
memo="暂无详细说明..."
Else
memo=trim(request.Form("memo"))
End if
dianhua=HTMLEncode(trim(request.Form("dianhua")))
CompPhone=trim(request.Form("CompPhone"))
email=HTMLEncode(trim(request.Form("email")))
fax=trim(request.Form("fax"))
mpname=trim(request.Form("mpname"))
leixing=trim(request.Form("leixing"))
qq=trim(request.Form("qq"))
weixin=trim(request.Form("weixin"))
dizhi=HTMLEncode(trim(request.Form("dizhi")))
youbian=trim(request.Form("youbian"))
sdays=trim(request.Form("days"))
Readinfo=trim(request.form("Readinfo"))
yz=xinxish
delpas=trim(request.Form("delpas"))
If delpas="" Then delpas=gen_key(5)
delpas=md5(delpas)

chkdick(biaoti) '判断用，重要要！
If Check_Str(biaoti)=True Then
call errdick(13)
response.end
end If

If Check_Str(name)=True Then
call errdick(47)
response.end
end If

If Check_Str(dizhi)=True Then
call errdick(48)
response.end
end If

arrkillword = split(killword,"|")'检查后台设定的标题过滤字符
for j = 0 to ubound(arrkillword)
if instr(biaoti,arrkillword(j))>0 then
call errdick(13)
response.end
exit for
end if
next

arrkillname = split(killname,"|")'检查后台设定的联系人过滤字符
for j = 0 to ubound(arrkillname)
if instr(name,arrkillname(j))>0 then
call errdick(47)
response.end
exit for
end if
next

arrkilldizhi = split(killname,"|")'检查后台设定的联系人地址过滤字符
for j = 0 to ubound(arrkilldizhi)
if instr(dizhi,arrkilldizhi(j))>0 then
call errdick(47)
response.end
exit for
end if
Next

arrkillmemo = split(killmemo,"|")'检查后台设定的信息详细内容过滤字符
for j = 0 to ubound(arrkillmemo)
if instr(memo,arrkillmemo(j))>0 then
call errdick(49)
response.end
exit for
end if
Next

if codekg1=1 then
if Trim(Request.Form("wenti"))=Empty Or trim(request.form("daan"))<>trim(request.form("wenti")) then
call errdick(44)
response.end
end If
end if

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

if biaoti="" Then
response.write "<li>系统保护！"
response.end
end If


if session("addxinxi")<>"" Or Request.Cookies("addxinxi")<>"" Then
  if DateDiff("s",session("addxinxi"),Now())<infosj Or DateDIff("s",Request.Cookies("addxinxi"),Now())<infosj then
  response.write "<li>系统保护：你提交数据太快，系统中止运行，请等待"&infosj&"秒钟！"
  response.end
  end if
end If

Call OpenConn
	city_one=Conn.Execute("select city from DNJY_city where twoid=0 and id="&city_oneid)(0)
	If city_twoid>0 Then city_two=Conn.Execute("select city from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid)(0)
	If city_threeid<>0 Then city_three=Conn.Execute("select city from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid)(0)
	type_one=Conn.Execute("select name from DNJY_type where twoid=0 and id="&type_oneid)(0)
	If type_twoid<>0 Then type_two=Conn.Execute("select name from DNJY_type where id="&type_oneid&" and threeid=0 and twoid="&type_twoid)(0)
	If type_threeid<>0 Then type_three=Conn.Execute("select name from DNJY_type where id="&type_oneid&" and twoid="&type_twoid&" and threeid="&type_threeid)(0)

Call CloseRs(rs)
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_info"
rs.open sql,conn,1,3
rs.addnew

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
rs("username")="游客"
rs("name")=name
rs("biaoti")=biaoti
rs("keywords")=keywords
rs("leixing")=leixing
If scjiage<>"" Then rs("scjiage")=scjiage
rs("jiage")=jiage
rs("wcolor")=wcolor
rs("ncolor")=ncolor
rs("mpname")=mpname
rs("fax")=fax
rs("memo")=memo
rs("email")=email
rs("dianhua")=dianhua
rs("CompPhone")=CompPhone
rs("fbts")=sdays
rs("qq")=qq
rs("weixin")=weixin
rs("dizhi")=dizhi
rs("youbian")=youbian
rs("Readinfo")=Readinfo
rs("addsj")=now()
rs("fbsj")=now()
rs("dqsj")= dateadd("d", sdays, now)
if onoff=0  then
rs("yz")=0
else
rs("yz")=xinxish
end if
userip=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If userip="" Then 
userip=Request.ServerVariables("REMOTE_ADDR")
end if
rs("ip")=userip
rs("delpas")=delpas
rs.update
session("addxinxi")=now()
Response.Cookies("addxinxi")=now()
Call CloseRs(rs)

set rs=server.createobject("ADODB.recordset")
sql = "select top 1 id from DNJY_info order by id "&DN_OrderType&""
rs.open sql,conn,1,1
id=rs("id")
Call CloseRs(rs)
Call CloseDB(conn)
session("info")=session("info")+1 '限制一个会员期内（20分钟）发布信息次数


if onoff>0 and xinxish>0 Then



If infoip=0 then
response.write "<script language=JavaScript>" & chr(13) & "alert('恭喜你，信息发布成功！您已连续发布"&session("info")&"条信息！');" &"window.location='index.asp?"&C_ID&"'" & "</script>"
Else
response.write "<script language=JavaScript>" & chr(13) & "alert('恭喜你，信息发布成功！您已连续发布"&session("info")&"条信息，超"&infoip&"条将被锁定IP无法访问！');" &"window.location='index.asp?"&C_ID&"'" & "</script>"
End If

Else
If infoip=0 Then
response.write "<script language=JavaScript>" & chr(13) & "alert('恭喜你，信息发布成功,待管理员审核后显示。您已连续发布"&session("info")&"条信息！');" &"window.location='index.asp?"&C_ID&"'" & "</script>"
else
response.write "<script language=JavaScript>" & chr(13) & "alert('恭喜你，信息发布成功,待管理员审核后显示。您已连续发布"&session("info")&"条信息，超"&infoip&"条将被锁定IP无法访问！');" &"window.location='index.asp?"&C_ID&"'" & "</script>"
end if
end if
%>
            </td>
          </tr>
          <tr>
            <td height="26" colspan="3" bgcolor="#FFFFFF"></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
    <td width="4" valign="top" bgcolor="#e6e6e6"></td>
  </tr>
  <tr>
    <td colspan="3"><div align="center">
        <!--#include file=footer.asp-->
    </div></td>
  </tr>
</table>      
</body>
</html>
