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
dim id,a,user_a,username
id=trim(request("id"))
username=request.cookies("dick")("username")
set rs = Server.CreateObject("ADODB.RecordSet")
if request("dick")="chk" then
call edit_a()
response.end
else
sql="select a from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,1
if rs.eof then
response.write "��������"
response.end
end if
user_a=rs("a")
if user_a=<0 then
response.write "�㻹û�е��ߣ����ȹ�����ߣ�"
cl
response.end
end if
rs.close
sql="select a from [DNJY_info] where username='"&username&"' and id="&cstr(id)
rs.open sql,conn,1,1
if rs.eof then
response.write "��������"
response.end
end if
a=rs("a")
rs.close
%>
<body topmargin="0">
<title>�޸ı�����ɫ</title>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
<tr>
<td width="100%" height="25">��ǰ����Ϣ<%if a="0" then%><font color="#FF0000">û��ʹ��</font><%else%><font color="#FF0000">ʹ����</font><%end if%><font color="#0000FF">�����ɫ����
</font></td>
</tr>
<tr>
<td width="100%" height="25">�����ڻ���<%=user_a%>��<font color="#0000FF">�����ɫ����</font></td>
</tr>
<tr>
<form action="?id=<%=id%>&dick=chk" method="POST">
<td width="100%" height="25">����������ӻ��޸���ѡ�������б����ɫ
                          <%if user_a>0 then%>
                          <select name="a">
                          <option style="COLOR: black" value="0" selected>��ʹ��
                          </option>
                          <option style="background: #000088" value="000088">
                          </option>
                          <option style="background: #0000ff" value="0000ff">
                          </option>
                          <option style="background: #008800" value="008800">
                          </option>
                          <option style="background: #008888" value="008888">
                          </option>
                          <option style="background: #0088ff" value="0088ff">
                          </option>
                          <option style="background: #00a010" value="00a010">
                          </option>
                          <option style="background: #1100ff" value="1100ff">
                          </option>
                          <option style="background: #111111" value="111111">
                          </option>
                          <option style="background: #333333" value="333333">
                          </option>
                          <option style="background: #50b000" value="50b000">
                          </option>
                          <option style="background: #880000" value="880000">
                          </option>
                          <option style="background: #8800ff" value="8800ff">
                          </option>
                          <option style="background: #888800" value="888800">
                          </option>
                          <option style="background: #888888" value="888888">
                          </option>
                          <option style="background: #8888ff" value="8888ff">
                          </option>
                          <option style="background: #aa00cc" value="aa00cc">
                          </option>
                          <option style="background: #aaaa00" value="aaaa00">
                          </option>
                          <option style="background: #ccaa00" value="ccaa00">
                          </option>
                          <option style="background: #ff0000" value="ff0000">
                          </option>
                          <option style="background: #ff0088" value="ff0088">
                          </option>
                          <option style="background: #ff00ff" value="ff00ff">
                          </option>
                          <option style="background: #ff8800" value="ff8800">
                          </option>
                          <option style="background: #ff0005" value="ff0005">
                          </option>
                          <option style="background: #ff88ff" value="ff88ff">
                          </option>
                          <option style="background: #ee0005" value="ee0005">
                          </option>
                          <option style="background: #ee01ff" value="ee01ff">
                          </option>
                          <option style="background: #3388aa" value="3388aa">
                          </option>
                          <option style="background: #000000" value="000000">
                          </option>
                          </select> <input type="submit" value="�ύ" name="B1">
                          <%else%><font color="#999999">û�е���</font>
                          <%end if%>
</td>
</form>
</tr>
</table>
<p>�˲�����Ҫ֧��1��<font color="#0000FF">�����ɫ����</font></p>
<%
end if
sub edit_a()
sql="select a from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,3
a=rs("a")
if a<1 then
response.write "û���㹻�ĵ��ߣ�"
cl
response.end
else
rs("a")=rs("a")-1
rs.update
end if
rs.close
sql="select a from [DNJY_info] where username='"&username&"' and id="&cstr(id)
rs.open sql,conn,1,3
rs("a")=trim(request("a"))
rs.update
response.write "�޸ı�����ɫ�ɹ���"
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
