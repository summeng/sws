<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="config.asp"-->
<!--#include file="Include/Function.asp"-->
<!--#include file="Include/label.asp"-->
<!--#include file=usercookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<%Call OpenConn
dim id,d,user_d,i,username
i=1
id=trim(request("id"))
username=request.cookies("dick")("username")
set rs = Server.CreateObject("ADODB.RecordSet")
if request("dick")="chk" then
call edit_d()
response.end
else
sql="select d from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,1
if rs.eof then
response.write "��������"
response.end
end if
user_d=rs("d")
if user_d=0 then
response.write "�㻹û�е��ߣ����ȹ�����ߣ�"
cl
response.end
end if
rs.close
%>
<body topmargin="0">
<title>��Ϣ�޸�-��֤����ʹ��</title>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
<tr>
<td width="100%" height="25">
<p align="left">�����ڻ���<%=user_d%>��<font color="#0000FF">ͨ����֤����</font></td>
</tr>
<tr>
<form action="?id=<%=id%>&dick=chk" method="POST">
<td width="100%" height="25">
<p align="left">���ϣ���ø���Ϣֱ����ҳ����ʾ��ѡ���Ƿ�ʹ��<font color="#0000FF">��֤����</font>��
<%if user_d>0 then%>&nbsp;&nbsp;&nbsp;
<select name="d">
<option selected value="0">��ʹ��</option>
<option value="1">ʹ��</option>
</select>&nbsp;&nbsp;                     
				  <input type="submit" value="�ύ" name="B1"><%else%>
<font color="#999999">û�е���</font>
<%end if%>
</td>
</form>
</tr>
</table>
<p align="left">�˲�����������Ӧ�ĵ�������ʹ��һ����Ҫһ������</p>
<%
end if
sub edit_d()
sql="select d from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,3
d=rs("d")
if trim(request("d"))="0" then
response.write "��ѡ���˲�ʹ�ã����������������"
cl
response.end
end if
if d=<0 then
response.write "û���㹻�ĵ��ߣ�"
cl
response.end
else
rs("d")=rs("d")-1
rs.update
end if
rs.close
sql="select yz,fbsj from [DNJY_info] where username='"&username&"' and id="&cstr(id)
rs.open sql,conn,1,3
rs("yz")=1
rs("fbsj")=now()
rs.update
response.write "ʹ�õ��߳ɹ��������Ϣ�Ѿ�ͨ������֤��"
If Htmlhome=1 Then Call s_HomeHtml()'����������ҳ
rs.close
cl
response.end
end sub
%>
<%sub cl()%>
<body onLoad="setTimeout(window.close, 1000)">
<%end Sub
set rs=Nothing
Call CloseDB(conn)%>
