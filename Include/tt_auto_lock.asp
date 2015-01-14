<!--#include file=err.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%dim checktext,rslock,ip,nn,userip
Call OpenConn
'总检查限制
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

'会员连续登陆出错X次，锁定IP
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

'管理员连续登陆后台出错X次，锁定IP
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

'使用“取回密码”时连续出错X次，锁定IP
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

'一个会话期内（20分钟）连续发布留言X次，锁定IP
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

'一个会话期内（20分钟）连续发布信息X次，锁定IP
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

'一个会话期内（20分钟）连续出现sql注入行为X次，锁定IP
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
