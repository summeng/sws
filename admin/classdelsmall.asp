<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Call OpenConn
dim smallclassid,smallclassname
smallclassid=trim(request("smallclassid"))
smallclassname=trim(request("smallclassname"))
if smallclassid<>"" then
	sql="delete from DNJY_SmallClass where smallclassid="&cint(smallclassid)&""
	conn.execute sql
	sql="delete from DNJY_News where smallclassname='" & smallclassname & "'"
	conn.execute sql
end if
Call CloseDB(conn)      
response.redirect "classmanage.asp"
%>




                                                              
                                                              
                                                              