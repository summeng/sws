<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=config.asp-->
<!--#include file=usercookies.asp-->
<!--#include file="Include/head.asp"-->
<!--#include file="Include/Function.asp"-->
<!--#include file="Include/label.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
dim hb_all,a,b,ct,d,hb,username
a=trim(request.form("a"))
b=trim(request.form("b"))
ct=trim(request.form("ct"))
d=trim(request.form("d"))
if request.form("a")=0 and request.form("b")=0 and request.form("ct")=0 and request.form("d")=0 then
response.write "<script language=JavaScript>" & chr(13) & "alert('��������');" &"window.location='user_props.asp?"&C_ID&"'" & "</script>"
response.end
end if
username=request.cookies("dick")("username")
hb_all=a*g_a+b*g_b+ct*g_c+d*g_d
Call OpenConn
set rs=server.createobject("adodb.recordset")
sql = "select hb,jf,a,b,c,d from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,3
if rs("hb")<hb_all then
response.write "<script language=JavaScript>" & chr(13) & "alert('���"&rmb_mc&"̫�٣��޷�������ô����ߣ�');" &"window.location='user_props.asp?"&C_ID&"'" & "</script>"
response.end
else
rs("hb")=int(rs("hb")-hb_all)
rs("a")=rs("a")+int(a)
rs("b")=rs("b")+int(b)
rs("c")=rs("c")+int(ct)
rs("d")=rs("d")+int(d)
rs("jf")=int(rs("jf")+hb_all/jf_4)
rs.update
end if
response.write "<script language=JavaScript>" & chr(13) & "alert('������߳ɹ�����Ļ���ͬʱ���ӣ�');" &"window.location='user.asp?"&C_ID&"'" & "</script>"
Call CloseRs(rs)
Call CloseDB(conn)
%>
