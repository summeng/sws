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
<%Call OpenConn
dim id,i,str2,dick,xxid,bdback
bdback = Request.ServerVariables("HTTP_REFERER")
dick=request("dick")
id=trim(request("selectedid"))
str2=split(id,",")
for i=0 to ubound(str2)
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_reply] where id="&trim(str2(i))
rs.open sql,conn,1,3

if dick="delyz" then
rs("hfyz")=0
Conn.Execute("Update DNJY_info Set hfcs=hfcs-1,hfsj="&DN_Now&" where id="&rs("xxid"))
else
rs("hfyz")=1
Conn.Execute("Update DNJY_info Set hfcs=hfcs+1,hfsj="&DN_Now&" where id="&rs("xxid"))
end if
rs.update
Next
response.redirect (bdback)
Call CloseDB(conn)
%>