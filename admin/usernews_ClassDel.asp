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

<%
dim ClassName,ClassID,sSavePathFileName,aSavePathFileName,id,i
ClassID=ReplaceUnsafe(Request.QueryString("ClassID"))
if ClassID<>"" then
	sql="delete From DNJY_userNewsClass where id=" & ClassID
	conn.Execute sql		
	
	'ɾ�����ͼƬ
	Set Rs = Server.CreateObject( "ADODB.Recordset" )
	Sql = "SELECT id,content,SavePathFileName from DNJY_userNews WHERE ClassID=" & ClassID
	Rs.Open Sql, Conn, 1, 3
	If Not Rs.Eof Then
		do while not Rs.EOF 
		sSavePathFileName = Rs("SavePathFileName")		
		id=Rs("id")
		if sSavePathFileName<>"" then
				' �Ѵ�"|"���ַ���תΪ����
			aSavePathFileName = Split(sSavePathFileName, "|")
				' ɾ��������ص��ļ������ļ�����
			For i = 0 To UBound(aSavePathFileName)
				' ��·���ļ���ɾ���ļ�
				Call DoDelFile(aSavePathFileName(i))
			Next
   		 end if
	' ɾ������
	Rs.delete
    Rs.update	
	Rs.movenext
	loop 
	Rs.Close
	set Rs=nothing	
	End If	
end if
conn.close
set conn=nothing  
response.redirect "usernews_ClassManage.asp"
%>



