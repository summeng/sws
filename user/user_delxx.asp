<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=../config.ASP-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="../Include/md5.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
dim objFSO,fileExt,sql1,rs1,sql2,delpas
id=request("id")
if request("key")="del" then
%>
<form method="POST" action="?key=delok&id=<%=id%>">
<p align="center">������ɾ�����룺<input type="text" name="delpas" size="15"> </p>
<p align="center"> <input type="submit" value="ȷ��ɾ��" name="B1"></p>
</form>
�ο�ɾ���Լ���Ϣͨ����������Ϣ���������������뿪��
<%end if

if request("key")="delok" then
id=trim(request("id"))
delpas=md5(request("delpas"))
Call OpenConn
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_info] where delpas='"&delpas&"' and id="&cstr(id)
rs.open sql,conn,1,1
 if not(rs.eof or rs.bof) then
'================ɾ�������ɵ�htm�ļ�
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.fileExists(Server.MapPath("../html/"&trim(id)&".html")) then
objFSO.DeleteFile(Server.MapPath("../html/"&trim(id)&".html"))
end if
set objfso=Nothing
'==================================================================
set rs=server.createobject("adodb.recordset")
sql="delete from [DNJY_info] where id="&cstr(id)'ɾ����Ϣ
rs.open sql,conn,1,3
sql1="delete from [DNJY_favorites] where scid='"&cstr(id)&"'"'ɾ���ղ�
rs.open sql1,conn,1,3
sql2="delete from [DNJY_reply] where xxid="&cstr(id)&""'ɾ���ظ�
rs.open sql2,conn,1,3
response.Write "<script language=javascript>alert('ɾ����Ϣ�ɹ�!');window.close();</script>"
response.end
Else
response.write "<script language=JavaScript>" & chr(13) & "alert('ɾ�����벻��ȷ��');" & "history.back()" & "</script>"
end If
end If
%>
