<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="Include/Function.asp"-->
<!--#include file="Include/label.asp"-->
<!--#include file="SqlIn.Asp"-->
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
dim id,b,user_b,i,username
i=1
id=trim(request("id"))
username=request.cookies("dick")("username")
set rs = Server.CreateObject("ADODB.RecordSet")
if request("dick")="chk" then
call edit_b()
response.end
else
sql="select b from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,1
if rs.eof then
response.write "��������"
response.end
end if
user_b=rs("b")
if user_b<1 then
response.write "�㻹û�е��ߣ����ȹ�����ߣ�"
cl
response.end
end if
rs.close
sql="select b from [DNJY_info] where username='"&username&"' and id="&cstr(id)
rs.open sql,conn,1,1
if rs.eof then
response.write "��������"
response.end
end if
b=rs("b")
rs.close
%>
<body topmargin="0">
<title>��Ϣ�޸�-�ö�����ʹ��</title>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
<tr>
<td width="100%" height="25">��ǰ����Ϣ<%if b>"0" then%>ʹ����<%=b%>��<%else%>û��ʹ��<%end if%><font color="#0000FF">��Ϣ�ö�����
</font></td>
</tr>
<tr>
<td width="100%" height="25">�����ڻ���<%=user_b%>��<font color="#0000FF">��Ϣ�ö�����</font></td>
</tr>
<tr>

<form action="?dick=chk" name="form" method="POST">
<td width="100%" height="25">�������������ѡ������
                  <%if user_b>0 then%>
                  <select name="b">
                  <option value="-<%=b%>">��ʹ��</option>
				  <%for i=1 to user_b%>
				  <option value="<%=i%>"><%=i%></option>
				  <%next%>
				  </select>
				  <input type="hidden" name="id" value="<%=id%>">                     
				  <input class="inputb" type="submit" value="�ύ" name="B1">
                  <%else%><font color="#999999">û�е���</font>
                  <%end if%>
				  
</td>
</form>
</tr>
</table>
<p>�˲�����������Ӧ�ĵ�������</p>
<%
end if
sub edit_b()
sql="select b from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,3
b=rs("b")
if b<1 then
response.write "û���㹻�ĵ��ߣ�"
cl
response.end
else
rs("b")=rs("b")-int(trim(request.form("b")))
rs.update
end if
rs.close
sql="select b from [DNJY_info] where username='"&username&"' and id="&cstr(id)
rs.open sql,conn,1,3
if rs("b")=0 then
rs("b")=trim(request.form("b"))
else
rs("b")=rs("b")+int(trim(request.form("b")))
end if
rs.update
response.write "�޸���Ϣ�ö����߳ɹ���"
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
