<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Dim gl,qxyz,str
Server.ScriptTimeout=50000
id=trim(request("selectedid"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" &"window.location='infolist.asp'" & "</script>"
response.end
end If
gl=trim(request("gl"))

If gl="jj" then
qxyz=trim(request("qxyz"))
str=split(id,",")
for i=0 to ubound(str)
set rs=server.createobject("adodb.recordset")
sql="select hfxx,hfje from [DNJY_info] where id="&trim(str(i))
rs.open sql,conn,1,3
if qxyz="no" then
rs("hfxx")=0
rs.update
Else
If rs("hfje")<=0 Then
response.write "<script language=JavaScript>" & chr(13) & "alert('����Ϣ���۽��Ϊ0����������Ϊ������Ϣ��');</script>"
else
rs("hfxx")=1
rs.update
end If
end If
Next
end If

response.Write "<script language=javascript>location.replace('infolist.asp?page="&request("page")&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&request("dick")&"&city_one="&request("city_one")&"&city_two="&request("city_two")&"&city_three="&request("city_three")&"');</script>"

%>