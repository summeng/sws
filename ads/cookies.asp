<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
if request.cookies("administrator")="" or request.cookies("admindick")<>"1" or request.cookies("domain")="" then 
response.redirect "../login.asp" 
end if
if Request.ServerVariables("SERVER_NAME")<>request.cookies("domain") then
response.redirect "../login.asp"
end If

sub checkmanage(str)'Ȩ���ж�
Call OpenConn
Dim mrs,manage
Set mrs = conn.Execute("select * from DNJY_admin where username='"&request.cookies("administrator")&"'")
if not (mrs.bof and mrs.eof) then
manage=mrs("manage")
if instr(manage,str)<=0 then
conn.close
set conn = nothing
Response.Write ("<script language=javascript>alert('��Ǹ����û�д��������Ȩ�ޣ�');history.go(-1);</script>")
response.end
end if
else 
conn.close
set conn = nothing
Response.Write ("<script language=javascript>alert('��Ǹ����û�е�½������ִ�д˲�����');history.go(-1);</script>")

response.end
end if
set mrs=Nothing
end sub%>
