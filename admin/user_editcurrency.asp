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
dim id
id=trim(request("id"))
if not isnumeric(id) or id="" then
response.write "<li>��������"
response.end
end if
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select jf,hb from [DNJY_user] where id="&cstr(id)
rs.open sql,conn,1,1
%>
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<title>�޸�<%=rmb_mc%></title>
<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="200" height="35">
    <form action="user_editcurrency_save.asp?id=<%=id%>" method="POST">
    <tr>
      <td height="10" colspan="2">Ŀǰ�˻�Աӵ�л��ֺ�<%=rmb_mc%>Ϊ��</td>
      </td>
    </tr>
    <tr>
      <td width="60" height="10">
      <p align="center">���֣�</td>
      <td width="121" height="10">
      <input type="text" name="jf" size="12" value="<%=rs("jf")%>"></td>
    </tr>
    <tr>
      <td width="60" height="20">
      <p align="center"><%=rmb_mc%>��</td>
      <td width="121" height="20">
      <input type="text" name="hb" size="12" value="<%=rs("hb")%>"></td>
    </tr>
    <tr>
      <td width="60" height="6"></td>
      <td width="121" height="6">
      <p align="center">
      <input type="submit" value="�ύ" name="B1" style="color: #FFFFFF; border: 1px solid #000000; background-color: #666666"></td>
    </tr>
    <tr>
      <td width="200" height="6" colspan="2">
      <p align="center"><font color="#FF0000">ע�⣺�������Ϊ���֡�����̫��û�����壬���������ķ���Դ�����鲻����5000����32767����</font></td>
    </tr>
    </form>
  </table>
  </center>
</div>
<%
Call CloseRs(rs)
Call CloseDB(conn)
%>
