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
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Աע��-���</title>
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
response.write "��ֹվ���ύ�����"
response.end 
end if 

if sms_kg=True then
If CInt(trim(request.form("captcha")))<>session("sms_code") Then
Call Alert ("�ֻ�����У������ϵͳ���ɵĲ�����","-1")
End If
End If

if sms_kg=True And trim(request.form("mobile"))<>session("mobile") Then
Call Alert ("�ύ���ֻ�������֤�ֻ��Ų���ͬһ���룡","-1")
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
If Check_Str(username)=True Then'�Ƿ���������ַ�
call errdick(33)
Call CloseDB(conn)
response.end
end If

if nothaveChinese(username)=false then'����ʹ��������
call errdick(3)
Call CloseDB(conn)
response.end
end If

if trim(request.form("username"))="ALL" Or trim(request.form("username"))="all" Or trim(request.form("username"))="����" then'���Ƶļ���ע����
call errdick(22)
Call CloseDB(conn)
response.end
end If

if codekg1=1 then'�ʴ���֤��
if Trim(Request.Form("wenti"))=Empty Or trim(request.form("daan"))<>trim(request.form("wenti")) then
call errdick(44)
Call CloseDB(conn)
response.end
end If
end if

if codekg2=1 then'������֤��
if trim(request.form("yzm"))<>Session("pSN") then
call errdick(39)
Session("pSN")=""
Call CloseDB(conn)
response.end
end if
end if

if codekg3=1 then'������֤��
if lcase(Request.Form("code"))<>lcase(Session("GetCode")) then
call errdick(40)
Session("GetCode")=""
Call CloseDB(conn)
response.end
end If
end if

if codekg4=1 then'������֤
ttdv=cint(trim(request.form("ttdv")))+1
if ttdv<>cint(datepart("w",date())) Then
call errdick(41)
Call CloseDB(conn)
response.end
end If
end If

if codekg5=1 then'��ʽ��֤
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
 call errdick(22)'�û����Ѿ�ע���
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
 call errdick(23)'�ʼ���ַ�Ѿ�ע���
Call CloseRs(rs2)
Call CloseDB(conn)
 response.end
 end If 
 end If 

userip=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If userip="" Then 
userip=Request.ServerVariables("REMOTE_ADDR")'��ȡIP��ַ
end If

'����ע�����================================================
Dim timess,AppealNum,AppealCount
timess=times*3600 
AppealNum=zcNum 'ͬһIP�޶�ʱ������������1�� 
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
response.write "<p>&nbsp;<br><center>"&times&"Сʱ������ע��"&AppealNum&"�Σ�" 
Call CloseDB(conn)
response.end 
End If
'����ע�����end================================================

If mailqr=1 then'���ϵͳ����Ϊ�ʼ���֤
'���ݵ�ǰ��ַ,����ȷ�ϵ�ַ
Dim addr0,addr1,addr
addr0 = Request.Servervariables("server_name") 
addr1 = Request.Servervariables("url") 
addr1=replace(addr1,"user_reg_save.asp","check.asp") 
addr="http://"&addr0&addr1

'�����ǰ���ʼ��������ݿ���,������ȷ����.����ʹ�ÿ��е������
If IsqrEmail(email)=FALSE Then
Dim savemail,CodeBit,Ranid
CodeBit=20'ȷ���������λ��
	savemail=TRUE
	Ranid=makerndid(CodeBit)
	'���������ظ�,���ٴ�����һ���µ������
	Do While IsqrRanid(Ranid)
		Ranid=makerndid(CodeBit)
	Loop
End If

'ȷ������

End if'���ϵͳ����Ϊ�ʼ���֤END

If mailqr=1 then'���ϵͳ����Ϊ�ʼ���֤��ע�����ϱ��浽��ʱ���ݿ⣬���򱣴浽��ʽ���ݱ�ʼ**********
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_usertemp] where username='"&username&"'"
rs.open sql,conn,1,3
if rs.eof or rs.bof then'�ж��û����Ƿ����
'����û���������
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
session("sms_code")=""'����ֻ���֤�����
session("mobile")=""'�����֤�ֻ���
else'����û�������
 if trim(rs("username"))=username then
 call errdick(22)'�û����Ѿ�ע���
 Call CloseDB(conn)
 response.end
 end if
 call errdick(24)'δ֪��½����
 Call CloseDB(conn)
 response.end
end If'�ж��û����Ƿ����END
else'���ϵͳ���ò��ʼ���֤��ע�����ϱ��浽��ʽ���ݱ���**********
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,3
if rs.eof or rs.bof then'�ж��û����Ƿ����
'����û���������
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
If recommend<>"" Then rs("tjname")=recommend'�Ƽ���
rs.update
	rs.close	
	set rs=nothing
session("sms_code")=""'����ֻ���֤�����
session("mobile")=""'�����֤�ֻ���
'���ϵͳ��ע���û����ͻ�ӭվ�ڶ��ŵķ�����
set rs3=server.createobject("adodb.recordset")
sql3="select top 1 username from [DNJY_admin]"
rs3.open sql3,conn,1,1
admin=rs3("username")
Call CloseRs(rs3)
Response.cookies("reg")("regkey")=""

else'����û�������
 if trim(rs("username"))=username then
 call errdick(22)'�û����Ѿ�ע���
 Call CloseDB(conn)
 response.end
 end if
 call errdick(24)'δ֪��½����
 Call CloseDB(conn)
 response.end
end If'�ж��û����Ƿ����END

If recommend<>"" Then'ͬʱ���Ƽ������ӻ��ֺͷ��Ͷ���֪ͨ
Dim rsjf
    conn.execute "UPDATE DNJY_user SET jf=jf+"&jf_5&" WHERE username='"&recommend&"'" 
	set rsjf=server.createobject("adodb.recordset")
	sql="select * from DNJY_Message"
	rsjf.open sql,conn,1,3
	rsjf.addnew
	  	rsjf("name")=recommend		
	  	rsjf("neirong")="��ϲ�����Ƽ�������"&username&"ע�᱾վ��Ա����վ���������˻���"&jf_5&"�֡�"
	  	rsjf("riqi")=now()
	  	rsjf("fname")=admin
	rsjf.update	
	rsjf.close	
	set rsjf=nothing	
End If

End if'���ϵͳ����Ϊ�ʼ���֤��ע�����ϱ��浽��ʱ���ݿ⣬���򱣴浽��ʽ���ݱ�END**********


If regmail=1 And email<>"" And mailqr=0 Then call mail()'����ע��֪ͨ�ʼ�
If regmail=1 And email<>"" And mailqr=1 Then call mailcheck()'����ע��ȷ���ʼ�%>
<body>
<table width="500"  border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#111111" bgcolor="#EEEEEE" style="border-collapse: collapse">
<tr><td height=30>Ŀǰλ�ã�<a href=index.asp?<%=C_ID%>>��ҳ</a> > ��Աע�������(���)</td></tr>
    <tr>
      <td width="768" valign="top">��
        <div align="center">
        <center>
            <table width="489" border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
            
              <tr> 
                <td width="16">
                ��</td>
                <td width="456" bgcolor="#FFFFFF">
                  <table width="454" height="120" border="1" cellpadding="3" cellspacing="3" bordercolor="#CCCCCC" class="greyfont">
                    
                    <tr> 
                      <td height="19" bgcolor="#F5F5F5" width="454">
<p align="center" style="line-height: 200%; margin-top: 0; margin-bottom: 0">��ϲ<%=name%>��<b><font color="#FF0000">ע��ɹ��ˣ�</b></td>
                    </tr>
                    
                    <tr> 
                      <td height="19" bgcolor="#F5F5F5" width="454">
<%If regmail=1 And mailqr=0 then%><li>
<p align="center">&nbsp;ϵͳ�ѽ�ע����Ϣ���͵���|<b><font color=""#ff0000""><%=email%></b>|���뱣�ܺ�����ע������</p>
<%End if%>
<%If mailqr=0 then%>
<table class="font_10_e_black" cellspacing="0" cellpadding="3" width="98%" align="center" border="0">
                        <!---------------------->
                      </table>
              
                        <table class="font_10_e_black" cellspacing="3" cellpadding="3" width="98%" align="center" border="0">
                          <form id="f1" name="thisForm" action="login_save.asp?userdl=ok&<%=C_ID%>" method="POST" >
                            <tr> 
                              <td width="48%"> 
                                <p align="center">&nbsp;��½�ʺ�:
                              </td>
                              <td width="52%"> 
                                <input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" class="form_e_10_black" id="username" maxlength="20" size="10" name="username2" value="<%=username%>">
                              </td>
                            </tr>
                            <tr> 
                              <td> 
                                <p align="center">&nbsp;��½����:
                              </td>
                              <td> 
                                <input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" class="form_e_10_black" id="password" maxlength="32" size="10" type="password"  value="<%=password%>" name="password">
                              </td>
                            </tr>
                            <tr> 
                              <td>��<input type="hidden" name="yzm" value="<%=Session("pSN")%>"></td>
                              <td> 
                                <input name="I3" type="submit" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;"  value="�����½"  border="0">
                              </td>
                            </tr>
                          </form>
                        </table>
						<%End if%>
<li>
<p align="center">&nbsp;�����Ƽ�����������վע�ᣬ���Ƽ�����ע��ɹ�������û���<%=jf_5%>�֡�<br>�뽫�ƹ���븴�Ʒ����������ѣ�http://<%=web%>/user_regagree.asp?tjname=<%=username%><br>Ҳ�ɵ���Ա���Ļ�ø�����ʽ���Ƽ����롣</p></li>
<%If regmail=1 And mailqr=1 then%><li>
<p align="center">&nbsp;ע�᱾վ��Ա��Ҫ�ʼ���֤��ϵͳ�ѽ���֤��Ϣ���͵������䣺|<b><font color=""#ff0000""><%=email%></b>|������ղ���<%=regyxq%>����ȷ�ϣ����ڽ�ɾ��ע����Ϣ����ע����ʱ�ʼ������ൽ�����ʼ��䣬��Ҫ��Html��ʽ����ʼ���������ȷ�ϣ�</p></li><%End if%>
                    </td>
                    </tr>
                  </table>
                </td>
                <td width="17">
                ��</td>
              </tr>
 
            </table>
          
        </center>
        </div>��</td>
    </tr>
</table>

<%sub mail()'����ע��֪ͨ�ʼ�
dim mailbody,Jmail
mailbody="��ϲ���Ϊ"&webname&"��http://"&web&"���𾴻�Ա��<br><br>��Ļ�Ա��½�ʺţ�"&username&"<br><br>�����ھͿ��Ե�¼�����ַ���������Ϣ�ˣ�http://"&web&""
'�ʼ�����
Dim HostName,E_Subject,Information,S_Type,C_M_Chk,e_info
Email=email
HostName=webname
E_Subject="��ϲ��ע��ɹ�������յ�½�ʺţ�"
Information=mailbody
S_Type=True'TrueΪHTML��FalseΪ�ı���ʽ
C_M_Chk=True'�Ƿ�Smtp��������֤�û���������
If ChkMail(Email)=true Then Call DNJYEmail(Email,HostName,E_Subject,Information,S_Type,C_M_Chk)
'�ʼ�����END
end Sub

sub mailcheck()'���ͻ�Ա��֤�ʼ�
dim mailbody,Jmail
Dim qylink
qylink=addr&"?ranid="&ranid&"&recommend="&recommend&"&mailname="&email
Dim MailFormat,emailbody,mailSubject,dengji
emailbody="���ã���л��ע�� ��"&webname&"����Ա����Ա��¼�ʺ�"&username&"����Ϊ��ȷ��������ݣ�������"&regyxq&"���ڵ�������ȷ�����ӣ�����ɾ��ע�����ϣ�<br>" & vbcrlf
emailbody=emailbody  &"�����ʼ���ϵͳ�Զ����ͣ��벻Ҫֱ�ӻظ���<br>" & vbcrlf
emailbody=emailbody  &"<font size=2>����������������������������������������������������������������������<br>" & vbcrlf
emailbody=emailbody  &"���������������ȷ�����ע����Ч��(ע��Ҫ��Html��ʽ����ʼ�����ȷ��)��<br>" & vbcrlf
emailbody=emailbody  &"<a href="&qylink&">���ȷ��</a><br><br>" & vbcrlf
emailbody=emailbody  &qylink& vbcrlf &"<br>������ܵ���ģ�������������ε�ַ���ƺ������������ַ���س�ȷ�ϾͿ����ˡ�<br>" & vbcrlf
emailbody=emailbody  &"����������������������������������������������������������������������<br>" & vbcrlf
emailbody=emailbody  &""&webname&"&nbsp;&nbsp;http://"&web&"<br>" & vbcrlf
emailbody=emailbody  &"����������������������������������������������������������������������</font><br>" & vbcrlf
mailbody=emailbody
mailSubject = "��Աע��ȷ����"
dengji=1 
'�ʼ�����
Dim HostName,E_Subject,Information,S_Type,C_M_Chk,e_info
Email=email
HostName=webname
E_Subject=mailSubject
Information=mailbody
S_Type=True'TrueΪHTML��FalseΪ�ı���ʽ
C_M_Chk=True'�Ƿ�Smtp��������֤�û���������
If ChkMail(Email)=true Then Call DNJYEmail(Email,HostName,E_Subject,Information,S_Type,C_M_Chk)
'�ʼ�����END
end Sub
Call CloseDB(conn)%>