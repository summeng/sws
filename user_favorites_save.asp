<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=usercookies.asp-->
<!--#include file="config.asp"-->
<!--#include file="Include/head.asp"-->
<!--#include file="Include/Function.asp"-->
<!--#include file="Include/label.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
dim str2,username
username=request.cookies("dick")("username")
id=trim(request.Form("selectedid"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" & "history.back()" & "</script>"
response.end
end if
str2=split(id,",")
Call OpenConn
set rs=server.createobject("adodb.recordset")
for i=0 to ubound(str2)
sql="delete from [DNJY_favorites] where username='"&username&"' and id="&cstr(str2(i))
rs.open sql,conn,1,3
next
response.write "<script language=JavaScript>" & chr(13) & "alert('ɾ���ղسɹ���');" & "window.location='user_favorites.asp?"&L_id&"'" & "</script>"
response.end
Call CloseRs(rs)
Call CloseDB(conn)
%>
