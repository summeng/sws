<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
dim id
id=trim(request("id"))
if not isnumeric(id) or id="" then
response.write "<li>��������"
response.end
end if
dim gbook2
gbook2=trim(request("memo"))
Call OpenConn
set rs=server.createobject("adodb.recordset")
sql = "select gbook2,hfsj,hf from DNJY_gbook where id="&cstr(id)
rs.open sql,conn,1,3
if request("chk")="1" then
savehf
end if
%>
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<title>�ظ�����</title>
<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="240" height="280">
    <form action="?id=<%=id%>&chk=1" method="POST">
    <tr>
      <td width="240"  colspan="2">
      <p align="center">�ظ�����</td>
    </tr>
    <tr>
      <td width="240" height="20" colspan="2">
      <p align="center">
                    <textarea rows="18" name="memo" cols="35"><%if request("chk")="2" then%><%=rs("gbook2")%><%end if%></textarea></td>
    </tr>
    <tr>
      <td width="60" ></td>
      <td width="240" >
      <p align="center">
      <input type="submit" value="�ύ" name="B1" style="color: #FFFFFF; border: 1px solid #000000; background-color: #666666"></td>
    </tr>
    </form>
  </table>
  </center>
</div>
<%
sub savehf()
rs("gbook2")=gbook2
rs("hfsj")=now()
rs("hf")=1
rs.update
Call CloseRs(rs)
response.write "<li>�ظ��ɹ���"
cl
Call CloseDB(conn)
response.end
end sub

sub cl()%>
<body onLoad="setTimeout(window.close, 2000)">
<%end sub%>
