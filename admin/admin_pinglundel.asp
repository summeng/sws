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
dim id,str2,i
Server.ScriptTimeout=50000
id=trim(request("ID"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" &"window.location='admin_comment.asp'" & "</script>"
response.end
end if
str2=split(id,",")
for i=0 to ubound(str2)'ѭ����ʼ
set rs=server.createobject("adodb.recordset")
rs.open "delete from DNJY_news_pinglun where pinglunid="&cstr(str2(i)),conn,3
set rs=Nothing
Next'ѭ������
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('�ɹ�ɾ����');" & chr(13)
        If request.querystring("pl")="yes" Then
        response.write "window.document.location.href='admin_Comment.asp';"&chr(13)
		else
		response.write "window.document.location.href='admin_article.asp';"&chr(13)
		End if
		response.write "</script>" & chr(13)
response.end
Call CloseDB(conn)%>