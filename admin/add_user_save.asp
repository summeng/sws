<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=../config.asp-->
<!--#include file=../Include/md5.asp-->
<!--#include file=../Include/mail.asp-->
<!--#include file=cookies.asp-->
<!--#include file=../Include/err.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("04")
dim username,password,email,maillist,userip
username=request("username")
password=request("password")
If password<>"" Then
password=md5(password)
End If

email=request("email")
maillist=request("maillist")
userip=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If userip="" Then 
userip=Request.ServerVariables("REMOTE_ADDR")
end If
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_user] where username='"&username&"' or (email='"&email&"' and email<>'')"
rs.open sql,conn,1,3
if rs.eof or rs.bof Then
rs.addnew
rs("username")=username
rs("useryz")=1
rs("mailyz")=1
If password<>"" Then rs("password")=password
rs("email")=email
rs("maillist")=maillist
rs("zcdata")=now()
rs("jftime")=now()
rs("vipdq")=now()
rs.update

response.write "<li>��ӻ�Ա�ɹ���֪ͨ��Ա��ǰ̨�����й����ϡ�"
dim mailbody
mailbody="����"&webname&"��ע���Ա��Ϣ����½�ʺţ�"&username&" <br>��½��ַΪ��<a target=_blank href="&web&">"&web&"</a><br>��½�ɹ�����������������ϣ�����������������Ϣ��<br><br>�����ʼ���ϵͳ�Զ����ͣ�����ظ��ʼ���"
'�ʼ�����

Dim HostName,E_Subject,Information,S_Type,C_M_Chk,e_info
Email=email
HostName=webname
E_Subject="����"&webname&"��ע����Ϣ"
Information=mailbody
S_Type=True
C_M_Chk=True
e_info=DNJYEmail(Email,HostName,E_Subject,Information,S_Type,C_M_Chk)
response.write "<br>�����ʼ����:"&e_info&""
Call CloseDB(conn)
response.end

'�ʼ�����END
Else
response.write "<li>�û����������Ѵ��ڣ�����ѡ��"
end If
Call CloseDB(conn)%>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
-->
</style>