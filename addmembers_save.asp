<!--#include file="conn.asp"-->
<!--#include file=top.asp-->
<!--#include file=usercookies.asp-->
<!--#include file=Include/mail.asp-->
 <!--#include file="Include/err.asp"-->
 <!--#include file=source.asp-->
 <!--#include file=Include/md5.asp-->
 <!--#include file="Include/RemoteImg_save.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%if not ChkPost() then 
response.write "��ֹվ���ύ�����"
response.end 
end if
if Instr(Request.ServerVariables("HTTP_REFERER"),""&adduserinfo&"")=0 then
response.redirect ""&adduserinfo&"?"&C_ID&"" 
end If
Dim webqq
webqq=qq
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
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="980" height="371" bgcolor="#FFFFFF">
    <tr>
      <td width="214" height="356" valign="top" bgcolor="#FFFFFF"><div align="center">
      <div align="center"><!--#include file=userleft.asp--></div></td>
      <td width="760" height="299" valign="top" align="center"><table width="760" height="329" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
        <tr>
          <td width="100%" valign="top" height="329">��
           
<%dim arrkillword,arrkillname,arrkillmemo,arrkilldizhi,j,Enable,name,ttdv
dim biaoti,scjiage,jiage,wcolor,ncolor,memo,a,b,ct,d,zcdata,dizhi,youbian,CompPhone,weixin,sdays,Amount,hfje,email,xxmpname,dianhua,leixin,fbsj,hfcs,mobile,userqq,biaotixs,sdmba,usernameid,namea,xxtp,delpas,cksj,tpname,T_SaveImg,AspJpeg,T_UploadDir
username=request.cookies("dick")("username")
biaoti=HTMLEncode(trim(request.form("biaoti")))
keywords=HTMLEncode(Replace_Text(trim(request.Form("keywords"))))
name=HTMLEncode(trim(request.form("name")))
email=HTMLEncode(trim(request.form("email")))
fax=trim(request.Form("fax"))
mpname=trim(request.Form("mpname"))
city_oneid=Strint(Request.Form("city_one"))
city_twoid=Strint(Request.Form("city_two"))
city_threeid=Strint(Request.Form("city_three"))
type_oneid=Strint(Request.Form("type_one"))
type_twoid=Strint(Request.Form("type_two"))
type_threeid=Strint(Request.Form("type_three"))
leixing=request.form("leixing")
dianhua=HTMLEncode(trim(request.form("dianhua")))
scjiage=trim(request.form("scjiage"))
jiage=trim(request.form("jiage"))
wcolor=trim(request.Form("wcolor"))
ncolor=trim(request.Form("ncolor"))
dizhi=HTMLEncode(trim(request.form("dizhi")))
youbian=trim(request.form("youbian"))
CompPhone=trim(request.form("CompPhone"))
qq=trim(request.form("qq"))
weixin=trim(request.Form("weixin"))
a=trim(request.form("a"))
b=trim(request.form("b"))
ct=trim(request.form("ct"))
d=trim(request.form("d"))
sdays=trim(request.form("days"))
Readinfo=trim(request.form("Readinfo"))
If trim(request.form("Readinfo"))="" Then Readinfo=0
hfje=CInt(request.form("hfje"))
If trim(request("memo"))="" Then
Call Alert ("����д��Ϣ����","-1")
ElseIf CheckStringLength(trim(request("memo")))>memonum Then
Call Alert ("��Ϣ����������"&memonum&"���ַ���","-1")
else
memo=trim(request.Form("memo"))
End If
tpname=request.form("tpname")
delpas=md5(gen_key(5))

chkdick(biaoti) '�ж��ã���Ҫ��
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

arrkillword = split(killword,"|")'����̨�趨�ı�������ַ�
for j = 0 to ubound(arrkillword)
if instr(biaoti,arrkillword(j))>0 then
call errdick(13)
response.end
exit for
end if
next

arrkillname = split(killname,"|")'����̨�趨����ϵ�˹����ַ�
for j = 0 to ubound(arrkillname)
if instr(name,arrkillname(j))>0 then
call errdick(47)
response.end
exit for
end if
next

arrkilldizhi = split(killname,"|")'����̨�趨����ϵ�˵�ַ�����ַ�
for j = 0 to ubound(arrkilldizhi)
if instr(dizhi,arrkilldizhi(j))>0 then
call errdick(47)
response.end
exit for
end if
Next

arrkillmemo = split(killmemo,"|")'����̨�趨����Ϣ��ϸ���ݹ����ַ�
for j = 0 to ubound(arrkillmemo)
if instr(memo,arrkillmemo(j))>0 then
call errdick(49)
response.end
exit for
end if
next		



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
ttdv=cint(trim(request.form("ttdv")))+1
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
response.write "<li>ϵͳ������"
response.end
end If


if session("addxinxi")<>"" Or Request.Cookies("addxinxi")<>"" Then
  if DateDiff("s",session("addxinxi"),Now())<infosj Or DateDIff("s",Request.Cookies("addxinxi"),Now())<infosj then
  response.write "<li>ϵͳ���������ύ����̫�죬ϵͳ��ֹ���У���ȴ�"&infosj&"���ӣ�"
  response.end
  end if
end If

set rs=server.createobject("adodb.recordset")
sql = "select * from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,1
vip=rs("vip")
zcdata=rs("zcdata")
Amount=rs("Amount")
Call CloseRs(rs)

  if DateDiff("s",zcdata,Now())<zcfbsj then
  response.write "<li>ϵͳ������ע�᲻������������Ϣ��ϵͳ��ֹ���У���ȴ�"&zcfbsj&"�����ٷ�����"
  response.end
  end If
If Amount<CInt(hfje) Then
  Response.Write ("<script language=javascript>alert('�����ʻ����㣬���ܷ���������Ϣ�����ֵ�ʻ����ٷ���!');history.go(-1);</script>")
  response.end
  end If



	city_one=Conn.Execute("select city from DNJY_city where twoid=0 and id="&city_oneid)(0)
	If city_twoid>0 Then city_two=Conn.Execute("select city from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid)(0)
	If city_threeid<>0 Then city_three=Conn.Execute("select city from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid)(0)
	type_one=Conn.Execute("select name from DNJY_type where twoid=0 and id="&type_oneid)(0)
	If type_twoid<>0 Then type_two=Conn.Execute("select name from DNJY_type where id="&type_oneid&" and threeid=0 and twoid="&type_twoid)(0)
	If type_threeid<>0 Then type_three=Conn.Execute("select name from DNJY_type where id="&type_oneid&" and twoid="&type_twoid&" and threeid="&type_threeid)(0)

set rs=server.createobject("adodb.recordset")
sql = "select a,b,c,d from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,3

if len(a)=6 then
  if rs("a")<1 then
  errdick(25)
  response.end
  end if
rs("a")=rs("a")-1
end if
if b="" or b="0" then
b=0
else
  if rs("b")<int(b) then
  errdick(26)
  response.end
  end if
rs("b")=rs("b")-int(b)
end if
'if ct="1" then '�����ϴ��ɹ���Ų���
'  if rs("c")<1 then
'  errdick(27)
'  response.end
'  end if
'rs("c")=rs("c")-1
'end if
if d="1"then
  if rs("d")<1 then
  errdick(28)
  response.end
  end if
rs("d")=rs("d")-1
end if
rs.update
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
rs("username")=username
rs("biaoti")=biaoti
rs("keywords")=keywords
rs("leixing")=leixing
If scjiage<>"" Then rs("scjiage")=scjiage
rs("jiage")=jiage
rs("wcolor")=wcolor
rs("ncolor")=ncolor
rs("mpname")=mpname
rs("fax")=fax
rs("dizhi")=dizhi
rs("youbian")=youbian
rs("qq")=qq
rs("weixin")=weixin
rs("memo")=memo
If tpname<>"" then
rs("tupian")=tpname
rs("c")=1
Else
rs("tupian")=0
rs("c")=0
End if
rs("name")=name
rs("email")=email
rs("dianhua")=dianhua
rs("CompPhone")=CompPhone
rs("hfje")=hfje
rs("Readinfo")=Readinfo
If hfje>0 then
rs("hfxx")=1
End if
if a="" then
rs("a")=0
else
rs("a")=a
end if
if b="" then
rs("b")=0
else
rs("b")=b
end if
'if request.form("ct")="" then '�����ϴ��ɹ���Ų���
'rs("c")=0
'else
'rs("c")=trim(request.form("ct"))
'end if
if onoff>0 and (d="1" or vip>0 Or vipno>0) then
rs("yz")=1
else
rs("yz")=0
end If
rs("fbts")=sdays
rs("addsj")=now()
rs("fbsj")=now()
rs("dqsj")= dateadd("d", sdays, now)
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
conn.execute "UPDATE DNJY_user SET xxts=xxts+1 WHERE username='"&username&"'" 'ͬʱ���û����ݿ�����һ����Ϣ��¼
conn.execute "UPDATE DNJY_user SET jf=jf+"&jf_2&" WHERE username='"&username&"'" 'ͬʱ���û����ݿ����ӻ���

if ct="1" then
uptu
end if
session("info")=session("info")+1 '����һ����Ա���ڣ�20���ӣ�������Ϣ����

'����Ǿ�����Ϣ�����û����Ϳۿ�֪ͨ����������
If hfje>0 Then
dim mailbody,Jmail,mailbiaoti,s1
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_Financial] "
rs.open sql,conn,1,3
rs.addnew
rs("username")=username
rs("aliname")=name
rs("ywje")=hfje
rs("ywlx")="֧��"
rs("czbz")="����������Ϣ�ۿ�"
rs("czz")=username
rs("admincl")=1
rs("addTime")=now()
rs.update
Call CloseRs(rs)
conn.execute "UPDATE DNJY_user SET Amount=Amount-"&CInt(hfje)&" WHERE username='"&username&"'" 'ͬʱ���û������ʻ����
conn.execute "UPDATE DNJY_user SET tAmount=tAmount+"&CInt(hfje)&" WHERE username='"&username&"'" 'ͬʱ���û����������ܽ��

mailbiaoti="���ѳɹ�����������Ϣ"
s1="�𾴵�"&username&"���ã�����"&now()&"��["&webname&"]�����˾�����Ϣ��"&biaoti&"������������"&sdays&"�죬�ܽ��"&trim(request("hfje"))&"Ԫ��ϵͳ���Զ��������ʻ��п۳���Ӧ����"&trim(request("hfje"))&"Ԫ���������¼��վ��http://"&web&"�����û����Ŀɿ���������ϸ�� ��ϵQQ��"&webqq&"������"&webname&"�ͷ��������ʼ���ϵͳ�Զ����ͣ�����ֱ�ӻظ���"
mailbody=""&s1&""
'�ʼ�����
Dim HostName,E_Subject,Information,S_Type,C_M_Chk,e_info
Email=email
HostName=webname
E_Subject=mailbiaoti
Information=mailbody
S_Type=True
C_M_Chk=True
If ChkMail(Email)=true Then Call DNJYEmail(Email,HostName,E_Subject,Information,S_Type,C_M_Chk)
'�ʼ�����END

End If

if onoff>0 and (vipno>0 or vip>0) Then


'������������������������������������������
If infoip=0 then
response.write "<p align=""center"">��ϲ������Ϣ�����ɹ���������������"&session("info")&"����Ϣ��</p>"
Else
response.write "<p align=""center"">��ϲ������Ϣ�����ɹ���������������"&session("info")&"����Ϣ����"&infoip&"����������IP�޷����ʣ�</p>"
End if
Else
If infoip=0 Then
response.write "<p align=""center"">��ϲ������Ϣ�����ɹ���������Ա��˺���ʾ��������������"&session("info")&"����Ϣ��</p>"
else
response.write "<p align=""center"">��ϲ������Ϣ�����ɹ���������Ա��˺���ʾ��������������"&session("info")&"����Ϣ����"&infoip&"����������IP�޷����ʣ�</p>"
end if
end if
%>
              <title>��Ϣ����</title>
              <%sub uptu()%>
              <meta http-equiv="Content-Language" content="zh-cn">
              <link rel="stylesheet" type="text/css" href="1.CSS">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="86%">
                <tr>
                  <td width="100%"><p align="center">��ʹ����<b><font color="#FF0000">��ͼ����</font></b>�����������ϴ�ͼƬ
                    ����gif,jpg,bmp,png��ʽ����Ƚ���500�����ڣ�</td>
                </tr>
              </table>
            <form name="form1" method="post" action="upfile.asp?username=<%=username%>&vip=<%=vip%>&id=<%=id%>&<%=C_ID%>" enctype="multipart/form-data">
                <input type="hidden" name="id" value="<%=id%>">
                <p align="center">
                  <input type="file" name="file1" value="" size="26">
                  &nbsp;
                  <input type="submit" value="�ύ" name="B1">
                </p>
            </form>
            <%end sub%>
          </td>
        </tr>
      </table></td>
      <td width="4" align="center" valign="top" bgcolor="#e6e6e6"></td>
    </tr>
    <tr>
      <td height="26" colspan="4"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>
<%Call CloseDB(conn)%>