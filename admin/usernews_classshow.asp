<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=../Include/Function.asp-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("04")%>

<%Dim id,sh,act
id =ReplaceUnsafe(Request.QueryString("id"))
sh=ReplaceUnsafe(request.QueryString("sh"))
act=ReplaceUnsafe(request.QueryString("act"))
Set rs=Server.CreateObject("ADODB.RecordSet")
sql="select * From DNJY_userNewsClass where id = "&id
rs.Open sql,conn,1,3
if act="dh" then
	rs("dhshow")=sh
end if
if act="lm" then
	rs("lmshow")=sh
end if
rs.update
Call CloseRs(rs)
response.Redirect "usernews_ClassManage.asp"

%>

