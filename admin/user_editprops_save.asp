<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<%
dim id,username,email,a,b,c,d
a=trim(request("a"))
b=trim(request("b"))
c=trim(request("c"))
d=trim(request("d"))
id=trim(request("id"))
if not isnumeric(a) or a="" then
response.write "<li>����ֻ��Ϊ���֣�"
cl
response.end
end if
if not isnumeric(b) or b="" then
response.write "<li>����ֻ��Ϊ���֣�"
cl
response.end
end if
if not isnumeric(c) or c="" then
response.write "<li>����ֻ��Ϊ���֣�"
cl
response.end
end if
if not isnumeric(d) or d="" then
response.write "<li>����ֻ��Ϊ���֣�"
cl
response.end
end if
if not isnumeric(id) or id="" then
response.write "<li>��������1��"
cl
response.end
end If
Call OpenConn
set rs=server.createobject("adodb.recordset")
sql="select a,b,c,d from [DNJY_user] where id="&id
rs.open sql,conn,1,3
if rs.eof or rs.bof then
response.write "<li>��������"
cl
response.end
end if
rs("a")=a
rs("b")=b
rs("c")=c
rs("d")=d
rs.update
Call CloseRs(rs)
Call CloseDB(conn)
response.write "<li>�޸ĵ��߳ɹ���"
call cl()
%>
<%sub cl()%>
<body onLoad="setTimeout(window.close, 2000)">
<%end sub%>
