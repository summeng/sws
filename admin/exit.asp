<%@language=vbscript codepage=936 %>
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
Response.cookies("administrator")=""
Response.cookies("admindick")=""
session("admin_dick")=""
Response.Redirect "index.asp"
%>
