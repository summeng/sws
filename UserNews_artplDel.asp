<!--#include file="conn.asp"-->
<!--#include file=top.asp-->
<!--#include file=usercookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Dim isse,ClassName,se_ClassName,selecttype,searchkeys,page,pages,oRs,sSql,a,sSavePathFileName,aSavePathFileName
isse=ReplaceUnsafe(request.QueryString("isse"))
ClassName=ReplaceUnsafe(Request.QueryString("ClassName"))
se_Classname=ReplaceUnsafe(Request.QueryString("se_Classname"))
searchkeys=ReplaceUnsafe(Request.QueryString("searchkeys"))
selecttype=ReplaceUnsafe(request.QueryString("selecttype"))
if ReplaceUnsafe(Request.QueryString("page"))="" then
page=1
else
page=ReplaceUnsafe(Request.QueryString("page"))
end if
Call OpenConn
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
Call CloseDB(conn)
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
end if

%>