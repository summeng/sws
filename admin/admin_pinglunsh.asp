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
<%Call OpenConn
dim referer
referer = Request.ServerVariables("HTTP_REFERER")
	set rs=server.createobject("adodb.recordset")
	rs.open "select * from DNJY_news_pinglun where pinglunid="&request("id")&"",conn,1,3
If  Request("action") = "sh"  Then
	rs("sh")=1
	else
If  Request("action") = "shno"  Then
	rs("sh")=0
end If
end if
rs.update
Call CloseRs(rs)
response.redirect (referer)
Call CloseDB(conn)
%>

