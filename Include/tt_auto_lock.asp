<!--#include file=err.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%dim checktext,rslock,ip,nn,userip
Call OpenConn
'�ܼ������
if lockip<>"0" then
set rslock = Server.CreateObject("ADODB.RecordSet")
sql="select ip from DNJY_config"
rslock.open sql,conn,3,3
ip=rslock("ip")
rslock.close
set rslock=nothing
if ip<>"" then
'ip=checktext(ip)
lockip=split(cstr(ip),"@")
for Nn=0 to UBound(lockip)
if instr(Request.serverVariables("REMOTE_ADDR"),lockip(nn))>0 then
errdick(46)
Call CloseDB(conn)
response.end
end if
next
end if
end If

'��Ա������½����X�Σ�����IP
if session("login_error")>=hydlip and hydlip<>0 then
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from DNJY_config"
rs.open sql,conn,1,3
userip=Request.serverVariables("REMOTE_ADDR")
if instr(rs("ip"),userip)=0 Then
If rs("ip")="" Then
rs("ip")=userip
else
rs("ip")=trim(rs("ip"))&"@"&userip
End If
End If
rs.update
Call CloseRs(rs)
Call CloseDB(conn)
errdick(35)
response.end
end if

'����Ա������½��̨����X�Σ�����IP
if session("dick_err")>=htdlip and htdlip<>0 then
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from DNJY_config"
rs.open sql,conn,1,3
userip=Request.serverVariables("REMOTE_ADDR")
if instr(rs("ip"),userip)=0 Then
If rs("ip")="" Then
rs("ip")=userip
else
rs("ip")=trim(rs("ip"))&"@"&userip
End If
End If
rs.update
Call CloseRs(rs)
Call CloseDB(conn)
errdick(35)
response.end
end If

'ʹ�á�ȡ�����롱ʱ��������X�Σ�����IP
if session("gpw_error")>=gpwip and gpwip<>0 then
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from DNJY_config"
rs.open sql,conn,1,3
userip=Request.serverVariables("REMOTE_ADDR")
if instr(rs("ip"),userip)=0 Then
If rs("ip")="" Then
rs("ip")=userip
else
rs("ip")=trim(rs("ip"))&"@"&userip
End If
End If
rs.update
Call CloseRs(rs)
Call CloseDB(conn)
errdick(35)
response.end
end if

'һ���Ự���ڣ�20���ӣ�������������X�Σ�����IP
if session("book_error")>bookip and bookip<>0 then
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from DNJY_config"
rs.open sql,conn,1,3
userip=Request.serverVariables("REMOTE_ADDR")
if instr(rs("ip"),userip)=0 Then
If rs("ip")="" Then
rs("ip")=userip
else
rs("ip")=trim(rs("ip"))&"@"&userip
End If
End If
rs.update
Call CloseRs(rs)
Call CloseDB(conn)
errdick(36)
response.end
end if

'һ���Ự���ڣ�20���ӣ�����������ϢX�Σ�����IP
if session("info")>infoip and infoip<>0 then
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from DNJY_config"
rs.open sql,conn,1,3
userip=Request.serverVariables("REMOTE_ADDR")
if instr(rs("ip"),userip)=0 Then
If rs("ip")="" Then
rs("ip")=userip
else
rs("ip")=trim(rs("ip"))&"@"&userip
End If
End If
rs.update
Call CloseRs(rs)
Call CloseDB(conn)
errdick(36)
response.end
end if

'һ���Ự���ڣ�20���ӣ���������sqlע����ΪX�Σ�����IP
if session("sql_err")>sqlip and sqlip<>0 Then
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from DNJY_config"
rs.open sql,conn,1,3
userip=Request.serverVariables("REMOTE_ADDR")
if instr(rs("ip"),userip)=0 Then
If rs("ip")="" Then
rs("ip")=userip
else
rs("ip")=trim(rs("ip"))&"@"&userip
End If
End If
rs.update
Call CloseRs(rs)
Call CloseDB(conn)
errdick(37)
response.end
end If

Call CloseDB(conn)%>
