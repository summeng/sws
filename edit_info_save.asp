<!--#include file="conn.asp"-->
<!--#include file="Include/err.asp"-->
<!--#include file=usercookies.asp-->
<!--#include file="config.asp"-->
<!--#include file="Include/head.asp"-->
<!--#include file="Include/Function.asp"-->
<!--#include file="Include/label.asp"-->
<!--#include file=Include/mail.asp-->
<!--#include file="Include/RemoteImg_save.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Dim vip,webqq,username,biaoti,scjiage,jiage,memo,userip,zt,CompPhone,youbian,dizhi,sdays,Amount,gqje,user_c,cc,name,email,xxmpname,dianhua,j,arrkillword,arrkillname,arrkilldizhi,arrkillmemo,leixin,fbsj,hfcs,mobile,userqq,biaotixs,sdmba,usernameid,namea,xxtp,gxkfto,jf,wcolor,ncolor,cksj,tpname,T_SaveImg,AspJpeg,T_UploadDir
webqq=qq
username=request.cookies("dick")("username")
city_oneid=strint(request.form("city_one"))
city_twoid=strint(request.form("city_two"))
city_threeid=strint(request.form("city_three"))
type_oneid=strint(request.form("type_one"))
type_twoid=strint(request.form("type_two"))
type_threeid=strint(request.form("type_three"))
biaoti=HTMLEncode(trim(request.form("biaoti")))
keywords=HTMLEncode(Replace_Text(trim(request.Form("keywords"))))
name=HTMLEncode(trim(request.form("name")))
email=HTMLEncode(trim(request.form("email")))
dianhua=HTMLEncode(trim(request.form("dianhua")))
CompPhone=trim(request.form("CompPhone"))
youbian=trim(request.form("youbian"))
dizhi=HTMLEncode(trim(request.form("dizhi")))
qq=trim(request.form("qq"))
leixing=trim(request.form("leixing"))
scjiage=trim(request.form("scjiage"))
jiage=trim(request.form("jiage"))
zt=trim(request.form("zt"))
sdays=trim(request.form("days"))
sdays=trim(request.form("days"))
Readinfo=trim(request.form("Readinfo"))
gxkfto=trim(request.form("gxkfto"))
If trim(request.form("Readinfo"))="" Then Readinfo=0
If trim(request("memo"))="" Then
Call Alert ("����д��Ϣ����","-1")
ElseIf CheckStringLength(trim(request("memo")))>memonum Then
Call Alert ("��Ϣ����������"&memonum&"���ַ���","-1")
else
memo=trim(request.Form("memo"))
End If
tpname=request.form("tpname")

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
Next
Call OpenConn
set rs=server.createobject("adodb.recordset")
rs.open "select city from DNJY_city where twoid=0 and threeid=0 and id="&city_oneid,conn,1,1
city_one=rs("city")
rs.close
if city_twoid>0 then
rs.open "select city from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid,conn,1,1
city_two=rs("city")
rs.close
end if
if city_twoid>0 and city_threeid>0 then
rs.open "select city from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid,conn,1,1
city_three=rs("city")
rs.close
end if
rs.open "select name from DNJY_type where twoid=0 and threeid=0 and id="&type_oneid,conn,1,1
type_one=rs("name")
rs.close
if type_twoid>0 then
rs.open "select name from DNJY_type where id="&type_oneid&" and threeid=0 and twoid="&type_twoid,conn,1,1
type_two=rs("name")
rs.close
end if
if type_twoid>0 and type_threeid>0 then
rs.open "select name from DNJY_type where id="&type_oneid&" and threeid="&type_threeid&" and twoid="&type_twoid,conn,1,1
type_three=rs("name")
rs.close
end If

set rs=server.createobject("adodb.recordset")
sql = "select * from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,1
vip=rs("vip")
Amount=rs("Amount")
user_c=rs("c")
jf=rs("jf")
Call CloseRs(rs)

If Amount<CInt(request("hfje")) Then
  Response.Write ("<script language=javascript>alert('�����ʻ����㣬���ܷ���������Ϣ�����ֵ�ʻ����ٷ���!');history.go(-1);</script>")
Call CloseDB(conn)
  response.end
  end If

If gxkfto<>"" then
If jf<CInt(gxkfto) Then
  Response.Write ("<script language=javascript>alert('�����ʻ����ֲ���"&gxkf&"�֣����ܸ�����Ϣ����׬ȡ���ֺ��ٸ�����Ϣ!');history.go(-1);</script>")
Call CloseDB(conn)
  response.End
  Else
  conn.execute "UPDATE DNJY_user SET jf=jf-"&gxkf&" WHERE username='"&username&"'" 'ͬʱ���û������ʻ�����
end If
end If



id=trim(request("id"))
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_info where username='"&username&"' and id="&cstr(id)
rs.open sql,conn,1,3
if rs.eof then
errdick(0)
Call CloseRs(rs)
Call CloseDB(conn)
response.end
end If
gqje=Round((now()-rs("fbsj"))*Round(rs("hfje")/rs("fbts"),2),2) '���������Ѿ��۽��
rs("city_oneid")=city_oneid
rs("city_one")=city_one
rs("city_two")=city_two
rs("city_twoid")=city_twoid
rs("city_three")=city_three
rs("city_threeid")=city_threeid
rs("type_oneid")=type_oneid
rs("type_one")=type_one
rs("type_twoid")=type_twoid
rs("type_two")=type_two
rs("type_threeid")=type_threeid
rs("type_three")=type_three
rs("biaoti")=biaoti
rs("keywords")=keywords
rs("leixing")=leixing
If scjiage<>"" Then rs("scjiage")=scjiage
rs("jiage")=jiage
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
rs("qq")=qq
rs("youbian")=youbian
rs("dizhi")=dizhi
rs("weixin")=trim(request("weixin"))
rs("fax")=trim(request("fax"))
rs("mpname")=trim(request("mpname"))
rs("wcolor")=trim(request("wcolor"))
rs("ncolor")=trim(request("ncolor"))
rs("fbsj")=now()
rs("zt")=zt
rs("fbts")=sdays
rs("Readinfo")=Readinfo
rs("hfje")=rs("hfje")-gqje+CInt(trim(request.form("hfje")))
If trim(request.form("hfje"))>0 Then
rs("hfxx")=1
End if
if onoff>0 and (vipno>0 or vip>0) then
rs("yz")=1
else
rs("yz")=0
end if
rs("dqsj")= dateadd("d", sdays, now)
rs.update
cc=rs("c")
Call CloseRs(rs)

Dim ServerURL,aa,objfso,htmout,http
if onoff>0 and (vipno>0 or vip>0) And trim(request.form("Readinfo"))=0 Then



Else
'================ɾ�������ɵ��ļ�
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.fileExists(Server.MapPath("/html/"&id&".html")) then
objFSO.DeleteFile(Server.MapPath("/html/"&id&".html"))
end if
set objfso=Nothing
'===============================
End if


'����Ǿ�����Ϣ�����û����Ϳۿ�֪ͨ����������
If trim(request("hfje"))>0 Then
dim mailbody,Jmail,mailbiaoti,s1

set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_Financial] "
rs.open sql,conn,1,3
rs.addnew
rs("username")=username
rs("aliname")=trim(request("name"))
rs("ywje")=trim(request("hfje"))
rs("ywlx")="֧��"
rs("czbz")="����������Ϣ�ۿ�"
rs("czz")=username
rs("admincl")=1
rs("addTime")=now()
rs.update
Call CloseRs(rs)
conn.execute "UPDATE DNJY_user SET Amount=Amount-"&CInt(trim(request("hfje")))&" WHERE username='"&username&"'" 'ͬʱ���û������ʻ����
conn.execute "UPDATE DNJY_user SET tAmount=tAmount+"&CInt(trim(request("hfje")))&" WHERE username='"&username&"'" 'ͬʱ���û����������ܽ��

mailbiaoti="���ѳɹ�����������Ϣ"
s1="�𾴵�"&username&"���ã�����"&now()&"��["&webname&"]�����˾�����Ϣ��"&biaoti&"������������"&sdays&"�죬�ܽ��"&trim(request("hfje"))&"Ԫ��ϵͳ���Զ��������ʻ��п۳���Ӧ����"&trim(request("hfje"))&"Ԫ���������¼��վ��http://"&web&"�����û����Ŀɿ���������ϸ�� ��ϵQQ��"&webqq&"������"&webname&"�ͷ��������ʼ���ϵͳ�Զ����ͣ�����ֱ�ӻظ���"
mailbody=""&s1&""
'�ʼ�����
Dim HostName,E_Subject,Information,S_Type,C_M_Chk,e_info
Email=trim(request("email"))
HostName=webname
E_Subject=mailbiaoti
Information=mailbody
S_Type=True
C_M_Chk=True
If ChkMail(Email)=true Then Call DNJYEmail(Email,HostName,E_Subject,Information,S_Type,C_M_Chk)
'�ʼ�����END
End IF
'������������������������������������������


if cc="1" then
uptu
else

if onoff>0 and (vipno>0 or vip>0) then
response.write "<p align=""center"">��ϲ�㣬��"&request.form("d")&"��<font color=ff0000>"&biaoti&"</font>�ɹ���</p>"
else
response.write "<p align=""center"">��ϲ�㣬�޸�<font color=ff0000>"&biaoti&"</font>�ɹ����ȴ���ˣ�</p>"
end If
response.write "<meta http-equiv=refresh content=""1;URL=user_xxgl.asp?"&C_ID&""">"
end If
%>
              <title>�ϴ�ͼƬ</title>
              <%sub uptu()%>
              <meta http-equiv="Content-Language" content="zh-cn">
              <link rel="stylesheet" type="text/css" href="1.CSS">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="86%">
                <tr>
                  <td width="100%"><p align="center">��Ϣ���޸ĳɹ�!<p align="center">��ԭ�����ϴ���ͼƬ�����Ҫ�޸�ͼƬ�����������ϴ�ͼƬ
                    ��������޸�ͼƬ����<a href="user_xxgl.asp?<%=C_ID%>">����</a></td>
                </tr>
              </table>
<%if user_c<=0 then
response.write "<center><p>�����ͼ����Ϊ"&user_c&"�����޷��ϴ�ͼƬ��Ҫ�ϴ�ͼƬ���ȹ�����ߣ�"
Call CloseDB(conn)
response.End
else%>          <form name="form1" method="post" action="upfile.asp?username=<%=username%>&id=<%=id%>&cc=<%if cc=1 then%>1<%else%>0<%end if%>&vip=<%=vip%>" enctype="multipart/form-data" >
                &nbsp;<input type="hidden" name="id" value="<%=id%>">
                <p align="center">
                  ����gif,jpg,bmp,png��ʽ����Ƚ���500�����ڣ�<input type="file" name="file1" value="" size="26">
                  &nbsp;
                  <input type="submit" value="�ύ" name="B1">
                </p>
            </form>

<%end if
end Sub
Call CloseDB(conn)%>