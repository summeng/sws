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
<%if request.Form("pldel")="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" & "history.back()" & "</script>"
response.end
end if
Dim isse,ClassName,se_ClassName,selecttype,searchkeys,page,pages,id,oRs,sSql,a,sSavePathFileName,Submit,ClassID
isse=ReplaceUnsafe(request.QueryString("isse"))
ClassName=ReplaceUnsafe(Request.QueryString("ClassName"))
ClassID=Request.Form("ClassID")
se_Classname=ReplaceUnsafe(Request.QueryString("se_Classname"))
searchkeys=ReplaceUnsafe(Request.QueryString("searchkeys"))
selecttype=ReplaceUnsafe(request.QueryString("selecttype"))
Submit=Request.Form("Submit")

if ReplaceUnsafe(Request.QueryString("page"))="" then
page=1
else
page=ReplaceUnsafe(Request.QueryString("page"))
end if
If Submit="�������" Then
for each a in request.Form("pldel")
	id=a
	Set oRs = Server.CreateObject( "ADODB.Recordset" )
	sSql = "SELECT * from DNJY_userNews WHERE id=" & a
	oRs.Open sSql, Conn, 1, 3
	If Not oRs.Eof Then		
	oRs("ifshow")=1		
	oRs.update		
	End If	
	oRs.Close		
next
conn.close
set conn=Nothing
response.write"<SCRIPT language=JavaScript>alert('��˳ɹ���');"
response.write"location.href ='UserNews_Modifylist.asp?ClassName="&ClassName&"&isse="&isse&"&searchkeys="&searchkeys&"&selecttype="&selecttype&"&se_ClassName="&se_ClassName&"&page="&page&"'</SCRIPT>"
else
for each a in request.Form("pldel")
'ɾ�����ͼƬ
	id=a
	Set oRs = Server.CreateObject( "ADODB.Recordset" )
	sSql = "SELECT * from DNJY_userNews WHERE id=" & a
	oRs.Open sSql, Conn, 1, 3
	If Not oRs.Eof Then		
		sSavePathFileName = oRs("SavePathFileName")	
		
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
		oRs.delete
		oRs.update		
	End If	
	oRs.Close		
next
conn.close
set conn=nothing
if isse=0 then
Response.Redirect "UserNews_Modifylist.asp?isse="&isse&"&ClassName="&ClassName&"&page="&page
end if
if isse=1 then
	if selecttype=2 then
			response.write"<SCRIPT language=JavaScript>alert('ɾ���ɹ���');"
			response.write"location.href ='UserNews_Search.asp'</SCRIPT>"
	end if
	if selecttype=1 then
	Response.Redirect "UserNews_Modifylist.asp?isse="&isse&"&searchkeys="&searchkeys&"&selecttype="&selecttype&"&se_ClassName="&se_ClassName&"&page="&page
	end if
end If
end if

%>