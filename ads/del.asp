<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../SqlIn.Asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("13")
Call OpenConn
dim id,path,objFSO,F_filename,F_jsname
if Request.QueryString("id")<> "" then

id = Request.QueryString("id")
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from [DNJY_ad] where id=" & id
rs.Open sql,conn,1,3
		'ɾ��Ӳ�����ͼƬ�ļ���js�ļ�
		F_filename=Mid(rs("ADSrc"),6)
		F_jsname="js/"&rs("adid")&"_"&rs("city_oneid")&"_"&rs("city_twoid")&"_"&rs("city_threeid")&".js"

        If Mid(F_filename,1,1)<>"/" And F_filename<>"" then
        'if F_filename<>"" then
		If objFSO.fileExists(Server.MapPath(F_filename)) then
		objFSO.DeleteFile(Server.MapPath(F_filename))
		end If
		end if
		if objFSO.fileExists(Server.MapPath(F_jsname)) then
		objFSO.DeleteFile(Server.MapPath(F_jsname))		
		end if
rs.Delete 
set rs=nothing
conn.Close
set conn=nothing
set objFSO=nothing
Response.Redirect ("list.asp")
else
Response.Write "��������"
end if
%>