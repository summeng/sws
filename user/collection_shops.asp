<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Call OpenConn
dim username,scid
if request.cookies("dick")("username")="" then 
response.write "<br>"
response.write "<li>����û�е�½��"
cl
response.end
end if
scid=trim(request("scid"))
if scid="" then
response.write "<li>��������"
cl
response.end
end if
username=trim(request.cookies("dick")("username"))

set rs=server.createobject("adodb.recordset")
sql="select * from DNJY_sc_shops where username='"&username&"' and scid='"&scid&"' "
rs.open sql,conn,1,3
if (rs.eof or rs.bof) then
rs.addnew
rs("username")=username
rs("scid")=scid
rs("name")=trim(request("name"))
rs("sdname")=trim(request("sdname"))
rs("mpname")=trim(request("mpname"))
rs("scsj")=now()
rs.update
response.write "<br>"
response.write "<li>�ղسɹ���"
cl
else
response.write "<br>"
response.write "<li>���Ѿ��ղع��ˣ�"
cl
response.end
end if
Call CloseRs(rs)
Call CloseDB(conn)
%>
<%sub cl()%>
<body onLoad="setTimeout(window.close, 2000)">
<%end sub%>
