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
set rs=server.createobject("adodb.recordset")
rs.open "delete from DNJY_wenda where id="&request.querystring("id"),conn,3
set rs=Nothing
Call CloseDB(conn)
response.write"<SCRIPT language=JavaScript>alert('����ʹ�ɾ���ɹ���');"
response.write"location.href=""admin_check.asp"";</SCRIPT>"
Response.End
%>

