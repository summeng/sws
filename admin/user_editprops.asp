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
end If
Call OpenConn
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select a,b,c,d from [DNJY_user] where id="&cstr(id)
rs.open sql,conn,1,1
%>
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<title>�޸ĵ�������</title>
<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#CCCCCC" width="180" height="35">
    <form action="user_editprops_save.asp?id=<%=id%>" method="POST">
    <tr>
      <td width="80" height="25">
      <p align="center"><font color="#FF0000">��ɫ���ߣ�</font></td>
      <td width="101" height="25">
      <input type="text" name="a" size="10" value="<%=rs("a")%>" maxlength="10">
      <font color="#FF0000">��</font></td>
    </tr>
    <tr>
      <td width="80" height="25">
      <p align="center"><font color="#FF0000">�ö����ߣ�</font></td>
      <td width="101" height="25">
      <input type="text" name="b" size="10" value="<%=rs("b")%>" maxlength="10">
      <font color="#FF0000">��</font></td>
    </tr>
    <tr>
      <td width="80" height="25">
      <p align="center"><font color="#FF0000">ͼƬ���ߣ�</font></td>
      <td width="101" height="25">
      <input type="text" name="c" size="10" value="<%=rs("c")%>" maxlength="10">
      <font color="#FF0000">��</font></td>
    </tr>
    <tr>
      <td width="80" height="25">
      <p align="center"><font color="#FF0000">��֤���ߣ�</font></td>
      <td width="101" height="25">
      <input type="text" name="d" size="10" value="<%=rs("d")%>" maxlength="10">
      <font color="#FF0000">��</font></td>
    </tr>
    <tr>
      <td width="80" height="6"></td>
      <td width="101" height="6">
      <p align="center">
      <input type="submit" value="�ύ" name="B1" style="color: #FFFFFF; border: 1px solid #000000; background-color: #666666"></td>
    </tr>
    <tr>
      <td width="181" height="6" colspan="2">
      <p align="center"><font color="#FF0000">ȫ������Ϊ����</font></td>
    </tr>
    </form>
  </table>
  </center>
</div>
<%
Call CloseRs(rs)
Call CloseDB(conn)
%>
