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
dim bigclassname
bigclassname=trim(request("bigclassname"))
if bigclassname<>"" then
	sql="delete from DNJY_bigClass where bigclassname='" & bigclassname & "'"
	conn.execute sql
	sql="delete from DNJY_SmallClass where bigclassname='" & bigclassname & "'"
	conn.execute sql
	sql="delete from DNJY_News where bigclassname='" & bigclassname & "'"
	conn.execute sql
end if
Call CloseDB(conn)    
response.redirect "classmanage.asp"
%>




                                                              
                                                              
                                                              