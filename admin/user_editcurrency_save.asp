<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
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
dim id,username,email,hb,jf
id=trim(request("id"))
hb=trim(request("hb"))
jf=trim(request("jf"))
if not isnumeric(jf) or jf="" then
response.write "<li>����ֻ��Ϊ���֣�"
cl
response.end
end if
if not isnumeric(hb) or hb="" then
response.write "<li>"&rmb_mc&"ֻ��Ϊ���֣�"
cl
response.end
end if
if not isnumeric(id) or id="" then
response.write "<li>��������"
cl
response.end
end If
Call OpenConn
set rs=server.createobject("adodb.recordset")
sql="select hb,jf from [DNJY_user] where id="&id
rs.open sql,conn,1,3
if rs.eof or rs.bof then
response.write "<li>��������"
cl
response.end
end if
rs("jf")=jf
rs("hb")=hb
rs.update
Call CloseRs(rs)
Call CloseDB(conn)
response.write "<li>�޸ĳɹ���"
call cl()
%>
<%sub cl()%>
<body onLoad="setTimeout(window.close, 2000)">
<%end sub%>
