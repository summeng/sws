<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="../Include/mail.asp"-->
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
dim username,nm,userip,biaoti,rsmes,email
biaoti=trim(request("biaoti"))
id=trim(request("id"))
nm=Request.form("nm")
username=request("username")
email=request("email")
city_oneid=request("city_oneid")
city_twoid=request("city_twoid")
city_threeid=request("city_threeid")
if Request.form("r1")="0" And trim(request("xxusername"))="�ο�" Then
response.write "<li>���û�δע�ᣬ����ֱ�ӻظ�����ѡ��'���������˵�����'��<a href='javascript:history.back(-1);'>����</a>"
response.end
end If

if Request.form("r1")="1" And (request.cookies("dick")("username")="" or request.cookies("dick")("domain")="" or request.cookies("dick")("id")="")   Then
response.write "<br>"
response.write "<li>���ʼ���ʽ��Ҫ�ȵ�¼������û�е�½��"
response.write "<meta http-equiv=refresh content=""2;URL=../login.asp?"&C_ID&""">"
response.end
end If
if Request.form("validate_codee")=""  Then
response.write "<li>��֤��û����д��<a href='javascript:history.back(-1);'>����</a>"
cl
response.end
end If

if not isnumeric(id) or id="" then
response.write "<li>��������"
cl
response.end
end if
if nm="on" Then
username="����"
call dick()
elseif (request.cookies("dick")("username")="" or request.cookies("dick")("domain")="" or request.cookies("dick")("id")="") And newspl=1 then 
response.write "<br>"
response.write "<li>����û�е�½��"
response.write "<meta http-equiv=refresh content=""2;URL=../login.asp?"&C_ID&""">"
response.End
elseif (request.cookies("dick")("username")="" or request.cookies("dick")("domain")="" or request.cookies("dick")("id")="") And newspl=2 then 
Response.Write ("<script language=javascript>alert('��ֹ������!');history.go(-1);</script>")
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
response.write "<li>�ظ�����û����д��"
cl
response.end
end If

if len(trim(request.form("neirong")))>200 then
response.write "<li>�ظ����ݲ��ܴ���200���ַ���"
cl
response.end
end If

if CInt(trim(request.form("validate_codee")))<>CInt(Session("ttpSN")) then
response.write "��֤�벻�ԣ�"
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

'����Ϣ�����߷��Ͷ���
	sql="select * from DNJY_Message"
	set rsmes=server.createobject("adodb.recordset")
	rsmes.open sql,conn,1,3
    rsmes.AddNew
  	rsmes("name")=trim(request("xxusername"))
	neirong="���˶�����������Ϣ��"&biaoti&"���ܸ���Ȥ�������˻ظ����뼰ʱ�鿴������鿴�����ظ��������ǹ���Աδ��˻����δͨ����"
	rsmes("neirong")=neirong
  	rsmes("riqi")=now()
 	rsmes("fname")=username
rsmes.Update
rsmes.close
set rsmes=Nothing

'����Ϣ�����߷����ʼ�
dim mailbody,Jmail
mailbody="���˶����ڡ�<a href='http://"&web&"'>"&webname&"</a>����������Ϣ��"&biaoti&"���ܸ���Ȥ�������˻ظ����뼰ʱ�鿴������鿴�����ظ��������ǹ���Աδ��˻����δͨ��������ϵ�ظ������䣺"&request("mymail")&"�������ʼ���ϵͳ�Զ����ͣ�����ֱ�ӻظ���"
'�ʼ�����
Dim HostName,E_Subject,Information,S_Type,C_M_Chk,e_info
Email=email
HostName=webname
E_Subject="���˶�����"&webname&"��������Ϣ��"&biaoti&"���ܸ���Ȥ"
Information=mailbody
S_Type=True
C_M_Chk=True
If ChkMail(Email)=true Then Call DNJYEmail(Email,HostName,E_Subject,Information,S_Type,C_M_Chk)
'�ʼ�����END

Set Jmail=Nothing



Conn.Execute("Update DNJY_info Set hfcs=hfcs+1 where id="&cstr(id))
Conn.Execute("Update DNJY_info Set hfsj="&DN_Now&" where id="&cstr(id))
If plsh=0 Then
response.write "<li>�ظ��ɹ���"
else
response.write "<li>�ظ��ɹ�������Ա��˺�ɼ���"
End if
cl
end sub
%>
<%sub cl()
response.write "<meta http-equiv=refresh content=""2;URL=../"&x_path("",id,"",Readinfo)&""">"
end Sub
Call CloseDB(conn)%>