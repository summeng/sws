<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=../config.asp-->
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
dim del_username,sql1,sql2,rs1,id,str2,i,alC_ID,str1,k,username,sql4,str4,a,sql3,adminrs,adminsql
alC_ID=trim(request("selectedid"))
If alC_ID="" Then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" &"window.location='users_List.asp'" & "</script>"
response.end
End if
set rs=server.createobject("adodb.recordset")
str1=split(alC_ID,",")
for k=0 to ubound(str1)
set adminrs=server.createobject("adodb.recordset")
adminsql="select username from [DNJY_admin] where username='"&trim(str1(k))&"'"
adminrs.open adminsql,conn,1,1
if not adminrs.eof Or trim(str1(k))="ttwlkj" then
response.write "<script language=JavaScript>" & chr(13) & "alert('ѡ��Ļ�ԱΪ��վ����ר���û���"&trim(str1(k))&"������ɾ����������ѡ��');" &"window.location='users_List.asp'" & "</script>"
adminrs.close
set adminrs=nothing
response.end
End if
sql="select id from [DNJY_info] where username='"&trim(str1(k))&"'"
rs.open sql,conn,1,1
do while not rs.eof
id=id +","&rs("id")&""
rs.movenext
loop
rs.close
str2=split(id,",")
     for i=1 to ubound(str2)
     sql="delete from [DNJY_info] where username='"&trim(str1(k))&"'"
     rs.open sql,conn,1,3
     sql1="delete from [DNJY_favorites] where scid='"&cstr(str2(i))&"' or username='"&trim(str1(k))&"' "
     rs.open sql1,conn,1,3
     sql2="delete from [DNJY_reply] where xxid="&cstr(str2(i))&" or username='"&trim(str1(k))&"' "
     rs.open sql2,conn,1,3
     Next     
sql3="delete from [DNJY_gbook] where username='"&trim(str1(k))&"' "
rs.open sql3,conn,1,3
sql2="delete from [DNJY_user] where username='"&trim(str1(k))&"' "
rs.open sql2,conn,1,3

response.write "<script language=JavaScript>" & chr(13) & "alert('ɾ���û��ɹ���');" & "window.location='users_List.asp?page="&request("page")&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&request("dick")&"&city_one="&request("city_one")&"&city_two="&request("city_two")&"&city_three="&request("city_three")&"'" & "</script>"

Next
adminrs.close
set adminrs=nothing
set rs=nothing
Call CloseDB(conn)
%>