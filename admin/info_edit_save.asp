<!--#include file="../conn.asp"-->
<!--#include file="../Include/err.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file=../Include/mail.asp-->
<!--#include file="../Include/RemoteImg_save.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Dim Amount,webqq,xxtp
webqq=qq
Call OpenConn
set rs=server.createobject("adodb.recordset")
sql = "select * from [DNJY_user] where username='"&request.form("uname")&"'"
rs.open sql,conn,1,1
if rs.eof Then
Amount=0
else
Amount=rs("Amount")
End if
Call CloseRs(rs)

If Amount<CInt(request("hfje")) Then
Call Alert ("'���ʻ��ʽ��ʻ����"&Amount&"Ԫ�����㾺����Ϣ������پ��۽����ֵ�ʻ����ٷ���!","-1")
  response.end
  end If
function chkdick(char) 
if instr(char,"'") then 
response.write "�����ַ���"
response.end 
end if
if instr(char,"|") then
response.write "�����ַ���"
response.end 
end if
if instr(char,"<") then
response.write "�����ַ���"
response.end 
end if
if instr(char,">") then
response.write "�����ַ���"
response.end 
end if
if instr(char,"&") then
response.write "�����ַ���"
response.end 
end if
if instr(char,"%") then
response.write "�����ַ���"
response.end 
end if
if instr(char,";") then
response.write "�����ַ���"
response.end 
end if
if instr(char,"$") then
response.write "�����ַ���"
response.end 
end if
end function
dim urlpage,biaoti,scjiage,jiage,memo,a,username,zt,yz,CompPhone,youbian,dizhi,dqsj,sdays,hfje,fsohtm,homeEliteImage,gqje,leixin,fbsj,userip,hfcs,dianhua,mobile,userqq,email,xxmpname,biaotixs,sdmba,usernameid,namea,wcolor,ncolor,cksj,tpname,T_SaveImg,AspJpeg,T_UploadDir

city_oneid=strint(request.form("city_one"))
city_twoid=strint(request.form("city_two"))
city_threeid=strint(request.form("city_three"))
type_oneid=strint(request.form("type_one"))
type_twoid=strint(request.form("type_two"))
type_threeid=strint(request.form("type_three"))

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
username=request.cookies("dick")("username")
biaoti=HTMLEncode(trim(request("biaoti")))
keywords=HTMLEncode(Replace_Text(trim(request("keywords"))))
scjiage=trim(request("scjiage"))
jiage=trim(request("jiage"))
a=trim(request("a"))
zt=trim(request("zt"))
yz=request("yz")
sdays=trim(request("days"))
If trim(request("memo"))="" Then
Call Alert ("����д��Ϣ����","-1")
ElseIf CheckStringLength(trim(request("memo")))>memonum Then
Call Alert ("��Ϣ����������"&memonum&"���ַ���","-1")
else
memo=trim(request.Form("memo"))
End If
tpname=request.form("tpname")



id=trim(request("id"))
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_info where id="&cstr(id)
rs.open sql,conn,1,3
if rs.eof then
errdick(0)
response.end
end If
If rs("hfje")<0 Then rs("hfje")=0
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
rs("leixing")=request("leixing")
rs("jiage")=jiage
If scjiage<>"" Then rs("scjiage")=scjiage
rs("memo")=memo
If tpname<>"" then
rs("tupian")=tpname
rs("c")=1
Else
rs("tupian")=0
rs("c")=0
End if
rs("a")=trim(request("a"))
rs("b")=trim(request("b"))
rs("name")=HTMLEncode(trim(request("name")))
rs("email")=HTMLEncode(trim(request("email")))
rs("dianhua")=HTMLEncode(trim(request("dianhua")))
rs("CompPhone")=trim(request("CompPhone"))
rs("qq")=trim(request("qq"))
rs("youbian")=trim(request("youbian"))
rs("dizhi")=HTMLEncode(trim(request("dizhi")))
rs("weixin")=trim(request("weixin"))
rs("fax")=trim(request("fax"))
rs("mpname")=trim(request("mpname"))
rs("wcolor")=trim(request("wcolor"))
rs("ncolor")=trim(request("ncolor"))
rs("homeEliteImage")=trim(request("homeEliteImage"))
rs("hfje")=rs("hfje")-gqje+CInt(trim(request("hfje")))
If trim(request("hfje"))>0 Or rs("hfje")>0 then
rs("hfxx")=1
Else
rs("hfxx")=0
End if
rs("dqsj")= dateadd("d", sdays, now)
rs("fbts")=sdays
rs("fbsj")=now()
rs("zt")=trim(request("zt"))
rs("Readinfo")=trim(request.form("Readinfo"))
rs("yz")=request("yz")
If request("yz")=1 then
rs("fsohtm")=1
Else
rs("fsohtm")=0
End if
rs.update
Call CloseRs(rs)


If request("yz")=1 And trim(request.form("Readinfo"))=0 Then
Dim objfso



Else
'================ɾ�������ɵ��ļ�
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.fileExists(Server.MapPath("../html/"&id&".html")) then
objFSO.DeleteFile(Server.MapPath("../html/"&id&".html"))
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
rs("username")=request.form("uname")
rs("aliname")=trim(request("name"))
rs("ywje")=trim(request("hfje"))
rs("ywlx")="֧��"
rs("czbz")="����������Ϣ�ۿ�"
rs("czz")=username
rs("admincl")=1
rs("addTime")=now()
rs.update
Call CloseRs(rs)
conn.execute "UPDATE DNJY_user SET Amount=Amount-"&CInt(trim(request("hfje")))&" WHERE username='"&request.form("uname")&"'" 'ͬʱ���û������ʻ����
conn.execute "UPDATE DNJY_user SET tAmount=tAmount+"&CInt(trim(request("hfje")))&" WHERE username='"&request.form("uname")&"'" 'ͬʱ���û����������ܽ��
mailbiaoti="�ѳɹ�����������Ϣ"
s1="�𾴵�"&request.form("uname")&"���ã�["&webname&"]��"&now()&"��["&webname&"]�����������˾�����Ϣ��"&biaoti&"������������"&sdays&"�죬�ܽ��"&trim(request("hfje"))&"Ԫ��ϵͳ���Զ��������ʻ��п۳���Ӧ����"&trim(request("hfje"))&"Ԫ���������¼��վ��http://"&web&"�����û����Ŀɿ���������ϸ�� ��ϵQQ��"&webqq&"������"&webname&"�ͷ��������ʼ���ϵͳ�Զ����ͣ�����ֱ�ӻظ���"
mailbody=""&s1&""

'�ʼ�����
Dim HostName,E_Subject,Information,S_Type,C_M_Chk,e_info
Email=trim(request("email"))
HostName=webname
E_Subject=mailbiaoti
Information=mailbody
S_Type=True
C_M_Chk=True
Call DNJYEmail(Email,HostName,E_Subject,Information,S_Type,C_M_Chk)
'�ʼ�����END

End IF
'������������������������������������������
Call CloseDB(conn)
response.write "<p align=""center"">�޸�<font color=ff0000>"&biaoti&"</font>�ɹ���</p>"
urlpage="info_edit_save.asp"
response.Write "<script language=javascript>location.replace('Substationrss.asp?rsssc=yes&urlpage="&urlpage&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&type_one="&type_oneid&"&type_two="&type_twoid&"&type_three="&type_threeid&"');</script>"
%>
