<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=config.ASP-->
<!--#include file=usercookies.asp-->
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
dim username,str2,tupian,objFSO,fileExt,sql1,rs1,sql2
username=request.cookies("dick")("username")
id=trim(request("selectedid"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" &"window.location='user_xxgl.asp?"&C_ID&"'" & "</script>"
response.end
end if
str2=split(id,",")
Call OpenConn
set rs=server.createobject("adodb.recordset")
for i=0 to ubound(str2)
sql="select c,tupian from [DNJY_info] where username='"&username&"' and id="&cstr(str2(i))
rs.open sql,conn,1,1
'On Error Resume Next
if rs("c")=1 then
   tupian=rs("tupian")
   fileExt=lcase(right(tupian,4))
   if fileEXT=".gif" or fileEXT=".jpg" then
   call deltu()
   end if
end if
rs.close
'================ɾ�������ɵ�htm�ļ�
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.fileExists(Server.MapPath("html/"&trim(str2(i))&".html")) then
objFSO.DeleteFile(Server.MapPath("html/"&trim(str2(i))&".html"))
end if
set objfso=Nothing
'==================================================================

sql="delete from [DNJY_info] where id="&cstr(str2(i))
rs.open sql,conn,1,3
sql1="delete from [DNJY_favorites] where scid='"&cstr(str2(i))&"' "
rs.open sql1,conn,1,3
sql2="delete from [DNJY_reply] where xxid="&cstr(str2(i))&" "
rs.open sql2,conn,1,3
conn.execute "UPDATE DNJY_user SET xxts=xxts-1 WHERE username='"&username&"'" 'ͬʱ���û����ݿ����һ����Ϣ��¼
conn.execute "UPDATE DNJY_user SET jf=jf-"&jf_2&" WHERE username='"&username&"'" 'ͬʱ���û����ݿ���ٻ���
Next
If Htmlhome=1 Then Call s_HomeHtml()'����������ҳ
response.write "<script language=JavaScript>" & chr(13) & "alert('ɾ����Ϣ�ɹ���');" &"window.location='user_xxgl.asp?"&C_ID&"'" & "</script>"
Call CloseRs(rs)
Call CloseDB(conn)
response.end
sub deltu()
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.fileExists(Server.MapPath("UploadFiles/infopic/"&tupian)) then
objFSO.DeleteFile(Server.MapPath("UploadFiles/infopic/"&tupian))
end if
set objfso=nothing
end sub
%>
