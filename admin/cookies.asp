<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%session.Timeout=60'���ú�̨ͣ��ʱ��Ϊ����
if request.cookies("administrator")="" or request.cookies("admindick")<>"1"  or request.cookies("domain")="" or session("admin_dick")="" then 
response.redirect "login.asp" 
end if
if Request.ServerVariables("SERVER_NAME")<>request.cookies("domain") then
response.redirect "login.asp"
end if

sub checkmanage(str)'Ȩ���ж�
Call OpenConn
Dim mrs,manage
Set mrs = conn.Execute("select * from DNJY_admin where username='"&request.cookies("administrator")&"'")
if not (mrs.bof and mrs.eof) then
manage=mrs("manage")
if instr(manage,str)<=0 then
mrs.Close:Set mrs=Nothing
Call CloseDB(conn)
Response.Write ("<script language=javascript>alert('��Ǹ����û�д��������Ȩ�ޣ�');history.go(-1);</script>")
response.end
end if
else 
mrs.Close:Set mrs=Nothing
Call CloseDB(conn)
Response.Write ("<script language=javascript>alert('��Ǹ����û�е�½������ִ�д˲�����');history.go(-1);</script>")
response.end
end if
end sub%>
