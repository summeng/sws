<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<%dim id,ServerURL,n
id=strint(request.querystring("id"))
ServerURL=CStr(Request.ServerVariables("SCRIPT_NAME"))
n=InStrRev(ServerURL,"/") '���ұߵ�һ���ַ������"_"��λ��,nΪ����ֵ
ServerURL=left(ServerURL, n)'��ʾ���������"n"���ַ�ǰ����ַ�,
ServerURL="http://"&Request.ServerVariables("SERVER_NAME")&""&ServerURL
Call OpenConn
if request("action")="Bad" Then'��ʾ����ʱ��
Dim rs9
set rs9=server.createobject("adodb.recordset")
sql = "select * from DNJY_Bad_info where sh=1 and infoid="&cstr(id)
rs9.open sql,conn,1,3
response.Write "document.write('"&rs9.recordcount&"');" & vbCrLf
rs9.close
set rs9=Nothing
end If
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_info where yz=1 and id="&cstr(id)
if jiaoyi=0 then
sql=sql&" and zt="&jiaoyi
end If
if overdue=0 Then
sql=sql&" and dqsj>"&DN_Now&""
end If
sql=sql&" and id="&cstr(id)
rs.open sql,conn,1,3
if request("action")="dqsj" then
Dim sj,fbsj
fbsj=rs("fbsj")
if rs("dqsj")<>"" Then
sj=DateDiff("d",now(),""&rs("dqsj")&"")
if sj>0 then
response.Write "document.write('<font color=#FF0000> ʣ��"&sj&"</b>��</font>');" & vbCrLf
elseif sj=0 then
response.Write "document.write('<font color=""#414141"">���յ���</font>');" & vbCrLf
elseif sj<0 then
response.Write "document.write('<font color=""#FF0000""> �Ѿ�����</font>');" & vbCrLf
end If
end If
end If

if request("action")="llcs" Then'��ʾ�������
response.Write "document.write('"&rs("llcs")&"');" & vbCrLf
rs("llcs")=rs("llcs")+1
rs.update
end If

if request("action")="hfcs" Then'��ʾ�ظ�����
response.Write "document.write('"&rs("hfcs")&"');" & vbCrLf
end If

if request("action")="zt" Then'��ʾ����״̬
dim zt
zt=rs("zt")
if zt=1 then
response.write "document.write('<font color=""#ff0000"">�Ѿ���ɽ���</font>');" & vbCrLf
elseif zt=0 then
response.write "document.write('<font color=""#0066FF"">��δ��ɽ���</font>');" & vbCrLf
end if
end If

if request("action")="username" And request.cookies("dick")("username")<>"" Then'��ʾ�ظ������û���
response.Write "document.write('<= <FONT color=blue><FONT style=""CURSOR: hand"" onclick=username.value="""&request.cookies("dick")("username")&""" color=#0000ff>"&request.cookies("dick")("username")&"</FONT></FONT>');" & vbCrLf
end If

if request("action")="email" And request.cookies("dick")("username")<>"" Then'��ʾ�ظ���������
Dim usermail
If request.cookies("dick")("username")<>"" Then usermail=conn.Execute("Select email From DNJY_user Where username='"&request.cookies("dick")("username")&"'")(0)
response.Write "document.write('<= <FONT color=blue><FONT style=""CURSOR: hand"" onclick=mymail.value="""&usermail&""" color=#0000ff>"&usermail&"</FONT></FONT>');" & vbCrLf
end If

if request("action")="usernumber" Then'��ʾ���˻�Ա
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_user"
rs.open sql,conn,1,1
response.write "document.write('���߾�����"&rs.recordcount&"λ');" & vbCrLf
End If

Call CloseRs(rs)
Call CloseDB(conn)%>
