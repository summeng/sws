<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("06")
Call OpenConn
dim rstp,rspl,str2,i,id,firstImageName,objFSO,rsuser
Server.ScriptTimeout=50000
id=trim(request("DelID"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" &"window.location='admin_article.asp'" & "</script>"
response.end
end if
str2=split(id,",")
set rs=server.createobject("adodb.recordset")
set rstp = server.CreateObject ("Adodb.recordset")
set rspl = server.CreateObject ("Adodb.recordset")
for i=0 to ubound(str2)'ѭ����ʼ
'===============ɾ�����ͼƬ===========================
set rstp = server.CreateObject ("Adodb.recordset")
sql="select firstImageName,username from [DNJY_News] where id="&cstr(str2(i))
rstp.open sql,conn,1,1
if Not IsNull(rstp("firstImageName")) Then
   if Not rstp.eof Then firstImageName=rstp("firstImageName")
   call deltu()  
end If
'===============================================
'===============�ۻ�Ա����===========================
If isnull(rstp("username")) then
set rsuser=server.createobject("adodb.recordset")
sql="select * from [DNJY_user] where username='"&rstp("username")&"'"
rsuser.open sql,conn,1,3    
if Not rsuser.eof Then
rsuser("jf")=rsuser("jf")-tgjf'�ۻ�Ա����
rsuser.update
Call CloseRs(rsuser)
end If
end if
'===============================================
'===============ɾ���������===========================
set rspl = server.CreateObject ("Adodb.recordset")
rspl.Open "select * from DNJY_news_pinglun where id="&cstr(str2(i)),conn,3,3
do while not rspl.eof
rspl.delete
rspl.movenext
Loop
'===============================================

sql="delete from [DNJY_News] where id="&cstr(str2(i))
rs.open sql,conn,1,3

rspl.close
set rspl=Nothing
rstp.close
set rstp=nothing
Next'ѭ������


response.write "<script language='javascript'>" & chr(13)
		response.write "alert('�ɹ�ɾ����');" & chr(13)
		response.write "window.document.location.href='admin_article.asp';"&chr(13)
		response.write "</script>" & chr(13)
response.end
Call CloseRs(rs)
rspl.close
set rspl=Nothing
rstp.close
set rstp=nothing

sub deltu()
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.fileExists(Server.MapPath("../"&firstImageName)) then
objFSO.DeleteFile(Server.MapPath("../"&firstImageName))
end if
set objfso=nothing
end sub
Call CloseDB(conn)%>
