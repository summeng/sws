<!--#include file="../conn.asp"-->
<!--#include file=../config.asp-->
<!--#include file=../Include/mail.asp-->

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
<title>���͵����ʼ�</title>
<meta http-equiv="Ŀ¼����" content="�ı�/html; �ַ���=gb2312">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>

</head>
<BODY>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
.style1 {
	color: #FFFFFF;
	font-weight: bold;
	font-size: 12px;
}
.style2 {font-size: 12px}
-->
</style>
<center>
<%Dim aJMail
Set aJMail=Server.CreateObject("JMail.Message")   
If aJMail Is Nothing Then  
response.write "<script LANGUAGE='javascript'>alert('��ܰС��ʾ����������֧��JMail.Message�������������װ�����������������ʼ�����!');history.go(-1);</script>"
Response.End
End If
If Request("mailto")<>"" And Request.form("mymail") = "" Then  
response.write "<script LANGUAGE='javascript'>alert('��ܰС��ʾ������д�������䣬����Է���ϵ��!');history.go(-1);</script>"
Response.End
End If
Call OpenConn
If trim(Request("email"))="" then
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select email from DNJY_user Where username='"&trim(Request("username"))&"'"
rs.open sql,conn,1,1
if rs.eof then
response.write "δ�鵽���û����䣡"
Call CloseRs(rs)
Call CloseDB(conn)
response.End
Else
email=rs("email")
End If
else
email=trim(Request("email"))
End If
mailbody=Request("mailbody")+""&email&""
If Request("Mobile")<>"" Then mailbody=mailbody+"���ֻ���"&Request("Mobile")&""
dim userid,rs1,tomail,mailsubject,mailbody,host,ErrStr,jmail,sitename,email,i
if Request.form("action")="send" then
call mailto()
Else
call mailedit()
End if
sub mailto()
If ChkMail(email) =False Then
response.Write "<script language=javascript>alert('�����������ʽ����!');window.close();</script>"
response.end
End If
If ChkMail(Request("mymail")) =False Then
response.Write "<script language=javascript>alert('���������ʽ����!');window.close();</script>"
response.end
End If
if Request("mailsubject") <> "" then 	
	mailsubject = Request.form("mailsubject") '�ż�����
	mailbody = mailbody+"�����������ʼ���["&webname&"-http://"&web&"]ϵͳ�Զ����ͣ�����ֱ�ӻظ�������ϵ�ʼ����������䣺"&Request("mymail")&"" '�ż�����
	host = ""	
	ErrStr = "ture"	

session("mcs")=session("mcs")+1
If session("mcs")>3 Then
response.Write "<script language=javascript>alert('�����Ŵ���̫���ˣ�����Ϣһ��!');window.close();</script>"
response.end
End If

'jmail�ʼ���������ʼ���ʼ�������ֵҪ��ȡ��
Set jmail = Server.CreateObject("JMail.Message") 
jmail.silent = true '����������󣬷���FALSE��TRUE��ֵ
jmail.Logging = true '�����ʼ���־
jmail.Charset = "gb2312"  '�ʼ������ֱ���Ϊ����
If Request("radiobutton") = "html" Then
jmail.ContentType = "text/html"    '�ʼ��ĸ�ʽΪHTML��ʽ
end if
jmail.MailServerUserName = mailname '����smtp��������֤��½�� ���ʾ����κ�һ���û���Email��ַ�� 
jmail.MailServerPassword = mailpass '����smtp��������֤���� ���û�Email�ʺŶ�Ӧ�����룩 
jmail.From = mailform '���������䣬��������������ʼ�������������뷢�ŷ�����ͬ���û������䣡 
jmail.FromName = webname '���������� 
jmail.AddRecipient email '�ռ���Email 
jmail.Subject = mailsubject '�ʼ��ı��� 
jmail.Body = mailbody '�ʼ�������
JMail.Priority = 5      '�ʼ��Ľ������� 
jmail.Send (mailsmtp) 'ִ���ʼ�����smtp��������ַ����ҵ�ʾֵ�ַ��
ErrStr =jmail.errormessage
jmail.close 
set jmail = nothing '�رն���
'�ʼ����ͽ���

if ErrStr = "" then
	Response.Write "��ϲ���ʼ����ͳɹ���"
'����������ͬʱд�ʼ���־��ʼ��������
Dim rslog
set rslog=server.createobject("adodb.recordset")
sql="select * from DNJY_log"
rslog.open sql,conn,1,3
rslog.addnew
rslog("username")=request.cookies("dick")("username")
rslog("outbox")=mailform
rslog("inbox")=email
rslog("title")=mailsubject
rslog("content")=mailbody
rslog("Sendtime")=now()
rslog("lock")=0
rslog.update
rslog.close
set rslog=nothing
'��������ͬʱд�ʼ���־END��������	
else	
	Response.Write "���ź����ʼ�δ�ܷ��ͣ�<br>�����ʼ�ϵͳ�������á��Ƿ��ж�����־��Ա."
end if	'���ͽ���жϽ���%>
<body onLoad="setTimeout(window.close, 2000)">
<%Response.end
else		     
response.write "<script language='javascript'>"
response.write "alert('�����ˣ��ʼ��������������д��');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end                                                                          
end if  '�Ƿ����жϽ���                                                                         
Call CloseDB(conn)
end Sub
sub mailedit()%> 
<table border="1" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%" cellpadding="0">
  <tr bgcolor="#BDBEDE">
    <td width="100%" height="28" colspan="8" align="center"><b><span class="style2">���͵����ʼ�</span></b></td>
  </tr>
</table>
  <form name="sendmail" action="user_mail.asp?mailto=yes" method="post" >
 <table border="1" width="99%" height="161"  background="images/greystrip.gif"  bordercolordark="#FFFFFF"  cellspacing="0">
      <tr>
    <td width="100%" colspan="2" height="20">
      <p ><font color="#ff0000"> ע�⣺�������йط��ɣ��Ͻ����ñ�ϵͳ����Υ����������Ϣ��</font></p>
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="2" height="20" background="images/mmto.gif">
      <p align="center"></p>
    </td>
  </tr>
  <tr>
    <td width="18%" height="26">�������䣺</td>
    <td width="82%" height="26"> <%=email%></td><input type="hidden" name="tomail" value="<%=email%>"></td>              
  </tr>                                                                   
  <tr>                                                                   
    <td width="14%" height="26">�ʼ����⣺</td>                                                                   
    <td width="86%" height="26">�Ҷ����ġ�<%=Request("mailzt")%>���ܸ���Ȥ<input type="hidden" name="mailsubject" value="�Ҷ����ġ�<%=Request("mailzt")%>���ܸ���Ȥ">
    </td>                                                                   
  </tr>    
  <tr>                                                                           
    <td width="14%" height="26">�ʼ����ݣ�</td>                                                                          
    <td width="86%" height="26">���ã��Ҷ����ڡ�<%=webname%>�������ġ�<%=Request("mailzt")%>���ܸ���Ȥ������ϵ�����䣺<input type="text" name="mymail" style="border:1px black solid" size="25"><input type="hidden" name="mailbody" value="���ã��Ҷ����ڡ�<%=webname%>�������ġ�<%=Request("mailzt")%>���ܸ���Ȥ������ϵ�����䣺"></td>                                                                          
  </tr> 
    <tr>                                                                   
    <td width="14%" height="26">�����ֻ���</td>                                                                   
    <td width="86%" height="26"><input type="text" name="Mobile" style="border:1px black solid" size="15">
    �����д����Է���ϵ��</td>                                                                   
  </tr>                                                                        
  <tr>                                                                          
    <td width="100%" height="26" colspan="2">                                                                          
      <p align="center"><input type="submit" name="send" value="����" style="border-style: solid; border-width: 1">                                                                           
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" name="Submit" value="ȡ��" style="border-style: solid; border-width: 1">     <input type=hidden name="action" value="send"> 
	  <input type="hidden" name="username" value="<%=trim(Request("username"))%>">
	  <input type="hidden" name="email" value="<%=trim(Request("email"))%>">
    </td>                                                                          
  </tr>                                                                          
 </table></form> 
 <%end Sub%>
</center>                                                                              
</body>                                                                              
</html>
<!--#include virtual="/Include/bottom_superadmin.asp" -->