<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file=Function.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%Call OpenConn
dim username,rsadId
username=trim(request("username"))
if Not nothaveChinese(UserName) Then
response.Write "<font color=#FF0000>��վ���Ʋ����ú���ע��!</font>"
Elseif Check_Str(UserName) then
response.Write "<font color=#FF0000>����Ҫע����û���"&username&"���б�վ��ֹע����ַ�!</font>"
Else
set rsadId=server.CreateObject("adodb.recordset")
rsadId.open"select username from DNJY_user where  username='"&username&"'",conn,1,1
if not rsadId.eof then
response.Write "<font color=#FF0000>��Ǹ���û���"&username&"�Ѵ��ڣ�����ѡ��</font>"
else
response.Write "<font color=#0000ff>��ϲ���û���"&username&"�����ڣ�����ע�ᡣ</font>"
end If
rsadId.close
set rsadId=Nothing
End If
Call CloseDB(conn)%>