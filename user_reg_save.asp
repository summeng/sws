<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="config.asp"-->
<!--#include file=Include/md5.asp-->
<!--#include file="Include/tt_auto_lock.asp" -->
<!--#include file="Include/filename.asp"-->
<!--#include file="Include/head.asp"-->
<!--#include file="Include/Function.asp"-->
<!--#include file="Include/label.asp"-->
<!--#include file="Include/err.asp"-->
<!--#include file="Include/ipt.asp"-->
<!--#include file="Include/mail.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>会员注册-完成</title>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}
-->
</style></head>

 
<%if not ChkPost() then 
response.write "禁止站外提交或访问"
response.end 
end if 

if sms_kg=True then
If CInt(trim(request.form("captcha")))<>session("sms_code") Then
Call Alert ("手机短信校验码与系统生成的不符！","-1")
End If
End If

if sms_kg=True And trim(request.form("mobile"))<>session("mobile") Then
Call Alert ("提交的手机号与验证手机号不是同一号码！","-1")
End If

if Instr(Request.ServerVariables("HTTP_REFERER"),""&regto&"")=0 then
response.redirect ""&reg&"?"&C_ID&"" 
end If

Call OpenConn
dim username,password,password3,problem,answer1,answer,name,idcard,Sex,dianhua,CompPhone,weixin,youbian,dizhi,mpname,http,email,maillist,ttdv,rsmes,rs3,sql3,rst,sqlt,admin,neirong,recommend
dim rs2,sql2
username=trim(request.form("username"))
password=request.form("password")
password3=md5(password)
problem=request.form("problem")
answer=md5(request.form("answer1"))
email=trim(request.form("email"))
recommend=trim(request.form("recommend"))
maillist=request.form("maillist")
name=trim(request.form("name"))
idcard=trim(request.form("idcard"))
Sex=trim(request.form("Sex"))
dianhua=trim(request.form("dianhua"))
CompPhone=trim(request.form("mobile"))
qq=trim(request.form("qq"))
weixin=trim(request("weixin"))
youbian=trim(request.form("youbian"))
dizhi=trim(request.form("dizhi"))
mpname=trim(request.form("mpname"))
fax=trim(request.form("fax"))
http=trim(request.form("http"))
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
If Check_Str(username)=True Then'是否包含敏感字符
call errdick(33)
Call CloseDB(conn)
response.end
end If

if nothaveChinese(username)=false then'不能使用中文名
call errdick(3)
Call CloseDB(conn)
response.end
end If

if trim(request.form("username"))="ALL" Or trim(request.form("username"))="all" Or trim(request.form("username"))="匿名" then'限制的几个注册名
call errdick(22)
Call CloseDB(conn)
response.end
end If

if codekg1=1 then'问答验证码
if Trim(Request.Form("wenti"))=Empty Or trim(request.form("daan"))<>trim(request.form("wenti")) then
call errdick(44)
Call CloseDB(conn)
response.end
end If
end if

if codekg2=1 then'数字验证码
if trim(request.form("yzm"))<>Session("pSN") then
call errdick(39)
Session("pSN")=""
Call CloseDB(conn)
response.end
end if
end if

if codekg3=1 then'文字验证码
if lcase(Request.Form("code"))<>lcase(Session("GetCode")) then
call errdick(40)
Session("GetCode")=""
Call CloseDB(conn)
response.end
end If
end if

if codekg4=1 then'星期验证
ttdv=cint(trim(request.form("ttdv")))+1
if ttdv<>cint(datepart("w",date())) Then
call errdick(41)
Call CloseDB(conn)
response.end
end If
end If

if codekg5=1 then'算式验证
if CInt(trim(request.form("validate_codee")))<>CInt(Session("ttpSN")) then
call errdick(45)
Session("ttpSN")=""
Call CloseDB(conn)
response.end
end if
end If

set rs2=server.createobject("adodb.recordset")
sql2="select username from [DNJY_user]  where username='"&username&"'"
rs2.open sql2,conn,1,1
set rst=server.createobject("adodb.recordset")
sqlt="select username from [DNJY_usertemp]  where username='"&username&"'"
rst.open sqlt,conn,1,1
if not rs2.eof Or not rst.eof then
 call errdick(22)'用户名已经注册过
Call CloseRs(rs2)
rst.close
set rst=Nothing
Call CloseDB(conn)
response.end
end If

If email<>"" then
set rs2=server.createobject("adodb.recordset")
sql2="select email from [DNJY_user]  where email='"&email&"'"
rs2.open sql2,conn,1,1
set rst=server.createobject("adodb.recordset")
sqlt="select email from [DNJY_usertemp]  where email='"&email&"'"
rst.open sqlt,conn,1,1
if not rs2.eof Or not rst.eof then
 call errdick(23)'邮件地址已经注册过
Call CloseRs(rs2)
Call CloseDB(conn)
 response.end
 end If 
 end If 

userip=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If userip="" Then 
userip=Request.ServerVariables("REMOTE_ADDR")'获取IP地址
end If

'限制注册次数================================================
Dim timess,AppealNum,AppealCount
timess=times*3600 
AppealNum=zcNum '同一IP限定时间内请求限制1次 
AppealCount=Request.Cookies("AppealCount") 
If AppealCount="" Then 
response.Cookies("AppealCount")=1 
AppealCount=1
response.cookies("AppealCount").expires=dateadd("s",""&timess&"",now()) 
Else 
response.Cookies("AppealCount")=AppealCount+1 
response.cookies("AppealCount").expires=dateadd("s",""&timess&"",now()) 
End If 
if int(AppealCount)>int(AppealNum) then 
response.write "<p>&nbsp;<br><center>"&times&"小时内限制注册"&AppealNum&"次！" 
Call CloseDB(conn)
response.end 
End If
'限制注册次数end================================================

If mailqr=1 then'如果系统设置为邮件验证
'根据当前网址,生成确认地址
Dim addr0,addr1,addr
addr0 = Request.Servervariables("server_name") 
addr1 = Request.Servervariables("url") 
addr1=replace(addr1,"user_reg_save.asp","check.asp") 
addr="http://"&addr0&addr1

'如果当前的邮件不在数据库中,则生成确认码.否则使用库中的随机码
If IsqrEmail(email)=FALSE Then
Dim savemail,CodeBit,Ranid
CodeBit=20'确认信随机码位数
	savemail=TRUE
	Ranid=makerndid(CodeBit)
	'如果随机码重复,则再次生成一个新的随机码
	Do While IsqrRanid(Ranid)
		Ranid=makerndid(CodeBit)
	Loop
End If

'确认连接

End if'如果系统设置为邮件验证END

If mailqr=1 then'如果系统设置为邮件验证，注册资料保存到临时数据库，否则保存到正式数据表开始**********
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_usertemp] where username='"&username&"'"
rs.open sql,conn,1,3
if rs.eof or rs.bof then'判断用户名是否存在
'如果用户名不存在
rs.addnew
rs("username")=username
rs("password")=password3
rs("problem")=problem
rs("answer")=answer
rs("name")=name
rs("email")=email
rs("maillist")=maillist
rs("idcard")=idcard
rs("userip")=userip
rs("Sex")=Sex
rs("dianhua")=dianhua
rs("CompPhone")=CompPhone
if youbian<>"" then
rs("youbian")=youbian
end if
rs("dizhi")=dizhi
rs("mpname")=mpname
rs("fax")=fax
rs("http")=http
rs("qq")=qq
rs("weixin")=weixin
rs("zcdata")=now()
rs("vipdq")=now()
rs("a")=z_a
rs("b")=z_b
rs("c")=z_c
rs("d")=z_d
rs("jf")=jf_1
rs("hb")=z_hb
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
rs("ranid")=ranid
If usersh=0 Then rs("useryz")=1
If mailqr=0 Then rs("mailyz")=1
rs.update
session("sms_code")=""'清空手机验证随机码
session("mobile")=""'清空验证手机号
else'如果用户名存在
 if trim(rs("username"))=username then
 call errdick(22)'用户名已经注册过
 Call CloseDB(conn)
 response.end
 end if
 call errdick(24)'未知登陆错误
 Call CloseDB(conn)
 response.end
end If'判断用户名是否存在END
else'如果系统设置不邮件验证，注册资料保存到正式数据表中**********
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,3
if rs.eof or rs.bof then'判断用户名是否存在
'如果用户名不存在
rs.addnew
rs("username")=username
rs("password")=password3
rs("problem")=problem
rs("answer")=answer
rs("name")=name
rs("email")=email
rs("maillist")=maillist
rs("idcard")=idcard
rs("userip")=userip
rs("Sex")=Sex
rs("dianhua")=dianhua
rs("CompPhone")=CompPhone
if youbian<>"" then
rs("youbian")=youbian
end if
rs("dizhi")=dizhi
rs("mpname")=mpname
rs("fax")=fax
rs("http")=http
rs("qq")=qq
rs("weixin")=weixin
rs("zcdata")=now()
rs("jftime")=now()
rs("vipdq")=now()
rs("a")=z_a
rs("b")=z_b
rs("c")=z_c
rs("d")=z_d
rs("jf")=jf_1
rs("hb")=z_hb
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
rs("ranid")=ranid
If usersh=0 Then rs("useryz")=1
If mailqr=0 Then rs("mailyz")=1
If recommend<>"" Then rs("tjname")=recommend'推荐人
rs.update
	rs.close	
	set rs=nothing
session("sms_code")=""'清空手机验证随机码
session("mobile")=""'清空验证手机号
'获得系统向注册用户发送欢迎站内短信的发信人
set rs3=server.createobject("adodb.recordset")
sql3="select top 1 username from [DNJY_admin]"
rs3.open sql3,conn,1,1
admin=rs3("username")
Call CloseRs(rs3)
Response.cookies("reg")("regkey")=""

else'如果用户名存在
 if trim(rs("username"))=username then
 call errdick(22)'用户名已经注册过
 Call CloseDB(conn)
 response.end
 end if
 call errdick(24)'未知登陆错误
 Call CloseDB(conn)
 response.end
end If'判断用户名是否存在END

If recommend<>"" Then'同时向推荐人增加积分和发送短信通知
Dim rsjf
    conn.execute "UPDATE DNJY_user SET jf=jf+"&jf_5&" WHERE username='"&recommend&"'" 
	set rsjf=server.createobject("adodb.recordset")
	sql="select * from DNJY_Message"
	rsjf.open sql,conn,1,3
	rsjf.addnew
	  	rsjf("name")=recommend		
	  	rsjf("neirong")="恭喜！您推荐的新人"&username&"注册本站会员，本站给您奖励了积分"&jf_5&"分。"
	  	rsjf("riqi")=now()
	  	rsjf("fname")=admin
	rsjf.update	
	rsjf.close	
	set rsjf=nothing	
End If

End if'如果系统设置为邮件验证，注册资料保存到临时数据库，否则保存到正式数据表END**********


If regmail=1 And email<>"" And mailqr=0 Then call mail()'发送注册通知邮件
If regmail=1 And email<>"" And mailqr=1 Then call mailcheck()'发送注册确认邮件%>
<body>
<table width="500"  border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#111111" bgcolor="#EEEEEE" style="border-collapse: collapse">
<tr><td height=30>目前位置：<a href=index.asp?<%=C_ID%>>首页</a> > 会员注册第三步(完成)</td></tr>
    <tr>
      <td width="768" valign="top">　
        <div align="center">
        <center>
            <table width="489" border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
            
              <tr> 
                <td width="16">
                　</td>
                <td width="456" bgcolor="#FFFFFF">
                  <table width="454" height="120" border="1" cellpadding="3" cellspacing="3" bordercolor="#CCCCCC" class="greyfont">
                    
                    <tr> 
                      <td height="19" bgcolor="#F5F5F5" width="454">
<p align="center" style="line-height: 200%; margin-top: 0; margin-bottom: 0">恭喜<%=name%>！<b><font color="#FF0000">注册成功了！</b></td>
                    </tr>
                    
                    <tr> 
                      <td height="19" bgcolor="#F5F5F5" width="454">
<%If regmail=1 And mailqr=0 then%><li>
<p align="center">&nbsp;系统已将注册信息发送到了|<b><font color=""#ff0000""><%=email%></b>|，请保管好您的注册资料</p>
<%End if%>
<%If mailqr=0 then%>
<table class="font_10_e_black" cellspacing="0" cellpadding="3" width="98%" align="center" border="0">
                        <!---------------------->
                      </table>
              
                        <table class="font_10_e_black" cellspacing="3" cellpadding="3" width="98%" align="center" border="0">
                          <form id="f1" name="thisForm" action="login_save.asp?userdl=ok&<%=C_ID%>" method="POST" >
                            <tr> 
                              <td width="48%"> 
                                <p align="center">&nbsp;登陆帐号:
                              </td>
                              <td width="52%"> 
                                <input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" class="form_e_10_black" id="username" maxlength="20" size="10" name="username2" value="<%=username%>">
                              </td>
                            </tr>
                            <tr> 
                              <td> 
                                <p align="center">&nbsp;登陆密码:
                              </td>
                              <td> 
                                <input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" class="form_e_10_black" id="password" maxlength="32" size="10" type="password"  value="<%=password%>" name="password">
                              </td>
                            </tr>
                            <tr> 
                              <td>　<input type="hidden" name="yzm" value="<%=Session("pSN")%>"></td>
                              <td> 
                                <input name="I3" type="submit" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;"  value="点击登陆"  border="0">
                              </td>
                            </tr>
                          </form>
                        </table>
						<%End if%>
<li>
<p align="center">&nbsp;您可推荐其它人来本站注册，所推荐的人注册成功您将获得积分<%=jf_5%>分。<br>请将推广代码复制发给您的朋友：http://<%=web%>/user_regagree.asp?tjname=<%=username%><br>也可到会员中心获得更多样式的推荐代码。</p></li>
<%If regmail=1 And mailqr=1 then%><li>
<p align="center">&nbsp;注册本站会员需要邮件验证，系统已将验证信息发送到了邮箱：|<b><font color=""#ff0000""><%=email%></b>|，请查收并在<%=regyxq%>天内确认，逾期将删除注册信息！（注意有时邮件被归类到垃圾邮件箱，并要用Html方式浏览邮件才能正常确认）</p></li><%End if%>
                    </td>
                    </tr>
                  </table>
                </td>
                <td width="17">
                　</td>
              </tr>
 
            </table>
          
        </center>
        </div>　</td>
    </tr>
</table>

<%sub mail()'发送注册通知邮件
dim mailbody,Jmail
mailbody="恭喜你成为"&webname&"－http://"&web&"的尊敬会员！<br><br>你的会员登陆帐号："&username&"<br><br>你现在就可以登录这个地址发布你的信息了：http://"&web&""
'邮件发送
Dim HostName,E_Subject,Information,S_Type,C_M_Chk,e_info
Email=email
HostName=webname
E_Subject="恭喜你注册成功，请查收登陆帐号！"
Information=mailbody
S_Type=True'True为HTML，False为文本格式
C_M_Chk=True'是否Smtp服务器验证用户名和密码
If ChkMail(Email)=true Then Call DNJYEmail(Email,HostName,E_Subject,Information,S_Type,C_M_Chk)
'邮件发送END
end Sub

sub mailcheck()'发送会员验证邮件
dim mailbody,Jmail
Dim qylink
qylink=addr&"?ranid="&ranid&"&recommend="&recommend&"&mailname="&email
Dim MailFormat,emailbody,mailSubject,dengji
emailbody="您好，感谢您注册 『"&webname&"』会员（会员登录帐号"&username&"），为了确认您的身份，请您在"&regyxq&"天内点击下面的确认链接，逾期删除注册资料！<br>" & vbcrlf
emailbody=emailbody  &"【本邮件由系统自动发送，请不要直接回复】<br>" & vbcrlf
emailbody=emailbody  &"<font size=2>＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝<br>" & vbcrlf
emailbody=emailbody  &"请点击下面的链接以确认你的注册有效性(注意要用Html方式浏览邮件才能确认)：<br>" & vbcrlf
emailbody=emailbody  &"<a href="&qylink&">点击确认</a><br><br>" & vbcrlf
emailbody=emailbody  &qylink& vbcrlf &"<br>如果不能点击的，请您把上面这段地址复制后添入浏览器地址栏回车确认就可以了。<br>" & vbcrlf
emailbody=emailbody  &"＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝<br>" & vbcrlf
emailbody=emailbody  &""&webname&"&nbsp;&nbsp;http://"&web&"<br>" & vbcrlf
emailbody=emailbody  &"＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝</font><br>" & vbcrlf
mailbody=emailbody
mailSubject = "会员注册确认信"
dengji=1 
'邮件发送
Dim HostName,E_Subject,Information,S_Type,C_M_Chk,e_info
Email=email
HostName=webname
E_Subject=mailSubject
Information=mailbody
S_Type=True'True为HTML，False为文本格式
C_M_Chk=True'是否Smtp服务器验证用户名和密码
If ChkMail(Email)=true Then Call DNJYEmail(Email,HostName,E_Subject,Information,S_Type,C_M_Chk)
'邮件发送END
end Sub
Call CloseDB(conn)%>