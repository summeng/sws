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
<%call checkmanage("14")
Call OpenConn
dim wenti,daan,id
wenti=request.form("wenti")
daan=request.form("daan")
if wenti="" or daan="" then
response.write"<SCRIPT language=JavaScript>alert('����ʹ𰸶�����Ϊ�գ�');"
response.write"javascript:history.go(-1)</SCRIPT>"
Response.End
end If
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from DNJY_wenda"
	rs.open sql,conn,1,2

rs.addnew
rs("wenti")=wenti
rs("daan")=daan
rs("xs")=1
rs.update
Call CloseRs(rs)
Call CloseDB(conn)
response.write"<SCRIPT language=JavaScript>alert('����ʹ���ӳɹ���');"
response.write"location.href=""admin_check.asp"";</SCRIPT>"
Response.End
%>