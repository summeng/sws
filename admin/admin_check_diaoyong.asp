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
dim id,xs
id=request("id")
xs=request("action")
Set rs=Server.CreateObject("ADODB.Recordset")
rs.open "select * from DNJY_wenda where id="&id,conn,1,3
If xs="wg" then
rs("xs")=1
Else
rs("xs")=0
End if
rs.update
Call CloseRs(rs)
Call CloseDB(conn)
response.write"<SCRIPT language=JavaScript>alert('ǰ̨��ʾ���óɹ���');"
response.write"location.href=""admin_check.asp"";</SCRIPT>"
Response.End
%>