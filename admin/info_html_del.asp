<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
dim str2,objFSO,username
username=request.cookies("dick")("username")
id=trim(request("selectedid"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" &"window.location='info_html.asp'" & "</script>"
response.end
end if
str2=split(id,",")
Call OpenConn
for i=0 to ubound(str2)

'================ɾ�������ɵ��ļ�
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.fileExists(Server.MapPath("../html/"&trim(str2(i))&".html")) then
objFSO.DeleteFile(Server.MapPath("../html/"&trim(str2(i))&".html"))
end if
set objfso=Nothing
'===============================
conn.execute "UPDATE DNJY_info SET fsohtm=0 WHERE id="&cstr(trim(str2(i)))&"" 'ͬʱ��עδ����htm�ļ�
Next
response.write "<script language=JavaScript>" & chr(13) & "alert('ɾ����Ϣ���htm�ļ��ɹ���');" &"window.location='"&request("urlpage")&"?rsssc=yes&city_oneid="&city_oneid&"&city_twoid="&city_twoid&"&city_threeid&"&city_threeid&"&page="&request("page")&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&request("dick")&"&city_one="&request("city_one")&"&city_two="&request("city_two")&"&city_three="&request("city_three")&"'" & "</script>"
response.end

Call CloseDB(conn)
%>
