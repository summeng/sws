<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Call OpenConn
dim mailname,mailpass,mailform,mailsmtp
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from DNJY_mail order by id "&DN_OrderType&""
rs.open sql,conn,3,3
if rs.eof then
response.write "��û���ʼ�ϵͳ���������ڡ���վϵͳ����-�ʼ�ϵͳ�������ϵͳ�ʼ����Ͳ�����"
Call CloseRs(rs)
Call CloseDB(conn)
response.End
End if
mailname=rs("mailname") '��¼�û���
mailpass=rs("mailpass") '��¼����
mailform=rs("mailform") '����������
mailsmtp=rs("mailsmtp") '�ʼ���������ַ
Call CloseRs(rs)
Call CloseDB(conn)

'****************************************************
'��������ChkMail
'�� �ã������ʽ���
'�� ����Email ----Email��ַ
'����ֵ��True��ȷ��False����
'ʹ��ʾ����
'If ChkMail("ls535427@2221262.com") = True Then
'Response.Write "��ʽ��ȷ"
'Else
'Response.Write "��ʽ����"
'End If
'****************************************************
Public Function ChkMail(ByVal Email)
Dim Rep,Pmail : ChkMail = True : Set Rep = New RegExp
Rep.Pattern = "([\.a-zA-Z0-9_-]){2,10}@([a-zA-Z0-9_-]){2,10}(\.([a-zA-Z0-9]){2,}){1,4}$"
Pmail = Rep.Test(Email) : Set Rep = Nothing
If Not Pmail Then ChkMail = False
End Function

'**********************************jmail���ź���************************************
'��������������Ƽ�www.ip126.com
'��������DNJYEmail() 
'���ã�����Jmail4.3�������E-Mail
'������
'Email:���ͣ��ַ��������ã�����E-Mail�ĵ�ַ��
'HostName�ʼ�����������
'E_Subject:���ͣ��ַ��������ã��ż����⡣
'Information:���ͣ��ַ��������ã��ż����ݡ�
'S_Type:���ͣ�����ֵ�����ã��Ƿ�ΪHtml��ʽ�ż���TrueΪHtml��ʽ��FalseΪ�ı���ʽ�� 
'C_M_Chk:���ͣ�����ֵ�����ã�Smtp�������Ƿ���Ҫ�����֤ ,True��Ҫ��֤��False����Ҫ��֤��
'������ͳɹ�������������True���򷵻�False 
'***********************************************************************************

Function DNJYEmail(Email,HostName,E_Subject,Information,S_Type,C_M_Chk)'��������ֱ��ǣ��ʼ����������䣬�ʼ�����������,�ʼ����⣬�ʼ����� ���ż���ʽ���Ƿ�Ҫ��֤
Dim C_Email,C_HostName,C_Smtp,C_M_User,C_M_Pass
C_Email=mailform'�����ߵ�����
C_HostName=HostName'�����ߵ�����
C_Smtp=mailsmtp'Smtp��������ַ
C_M_User=mailname'���Smtp��������Ҫ��֤��ݣ��������û���
C_M_Pass=mailpass'����������


Dim DNJY_Jmail
Err.Clear
On Error Resume Next
If Email="" Or Information="" Or E_Subject="" Then
DNJYEmail="<br><li>����������䡢�ż����⡢�ż������Ƿ�Ϊ�գ�</li>" 
Exit Function
End If
set DNJY_Jmail=Server.CreateObject("Jmail.message")'���������ʼ��Ķ���
if err then 
DNJYEmail= "<br><li>������û�а�װJMail������޷������ʼ���</li>" 
err.clear 
exit function 
end if 
DNJY_Jmail.silent = true '����������󣬷���FALSE��TRUE��ֵ
DNJY_Jmail.Charset="gb2312" '�ʼ������ֱ���Ϊ����,��Ҫ�˾���ܳ�������
DNJY_Jmail.Logging=True'�����ʼ���־ 
DNJY_Jmail.From=C_Email '�����˵�E-MAIL��ַ
DNJY_Jmail.Fromname=C_HostName'�����ߵ�����
DNJY_Jmail.addrecipient Email'�ʼ��ռ��˵ĵ�ַ
DNJY_Jmail.subject=E_Subject'�ż�����
If S_Type=False Then'�ı���HTML��ʽ
DNJY_Jmail.appendtext Information
Else
DNJY_Jmail.AppendHtml Information
End If
DNJY_Jmail.maildomain=C_Smtp'Smtp��������ַ
If C_M_Chk=True Then'Smtp��������֤�û���������
DNJY_Jmail.mailserverusername=C_M_User'��¼�ʼ�������������û���
DNJY_Jmail.mailserverpassword=C_M_Pass'��¼�ʼ����������������
End If
DNJY_Jmail.Priority = 1'�ʼ��Ľ�������1 Ϊ��죬5 Ϊ������ 3 ΪĬ��ֵ
DNJY_Jmail.send(C_Smtp)'ִ���ʼ����ͣ�ͨ���ʼ���������ַ��
If Err.Number<>0 Then'�ж��Ƿ��ͳɹ�
DNJYEmail="<li>�ʼ�����ʧ�ܣ������ʼ�����ϵͳ�����������������Ƿ���ȷ��"'ʧ�ܷ���
Else
DNJYEmail="<li>�ʼ����ͳɹ���"'�ɹ�����
'����������ͬʱд�ʼ���־��ʼ��������
Dim rslog
set rslog=server.createobject("adodb.recordset")
sql="select * from DNJY_log"
rslog.open sql,conn,1,3
rslog.addnew
rslog("username")=C_HostName
rslog("outbox")=C_Email
rslog("inbox")=email
rslog("title")=E_Subject
rslog("content")=Information
rslog("Sendtime")=now()
rslog("lock")=0
rslog.update
rslog.close
set rslog=nothing
'��������ͬʱд�ʼ���־END��������	
End If
Err.Clear
Set DNJY_Jmail=nothing'���ٶ���
End Function %>                    