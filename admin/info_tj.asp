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
<%Call OpenConn
dim id,tj,checked
id=trim(request("id"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" &"window.location='infolist.asp'" & "</script>"
response.end
end if
if not isnumeric(id) or id="" then
response.write "<li>��������"
response.end
end if
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select tj from [DNJY_info] where id="&cstr(id)
rs.open sql,conn,1,1
%>
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<title>�޸�״̬</title>
<style type="text/css">
<!--
.style1 {
	color: #FF0000;
	font-weight: bold;
}
-->
</style>
<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
  <br>
  <table width="206" height="29" border="1" cellpadding="0" cellspacing="0" bordercolor="#111111" bgcolor="#CCCCCC" style="border-collapse: collapse">
    <form action="info_tj_save.asp?id=<%=id%>" method="POST">
    <tr>
      <td height="35" colspan="2"><span class="style1"> ����Ϣ�Ƽ���</span></td>
    </tr>
    <tr>
      <td width="96" height="10">
      <p align="center">�����Ƽ�����</td>
      <td width="104" height="30">
        <div align="center">
          <label>
          <select name="tj" id="tj">
            <option value="0" <%if rs("tj")="0" then%>selected<%end if%>>�����Ƽ�</option>
            <option value="1" <%if rs("tj")="1" then%>selected<%end if%>>��Ϣ�Ƽ�</option>
            <option value="2" <%if rs("tj")="2" then%>selected<%end if%>>ͼƬ�Ƽ�</option>
          </select>
          </label>
        </div></td>
    </tr>
    <tr>
      <td height="30" colspan="2">
        <p align="center">
          <input type="submit" value="�� ��" name="B1" style="color: #FFFFFF; border: 1px solid #000000; background-color: #666666">
        </td>
      </tr>
    </form>
  </table>
  </center>
</div>
<%
Call CloseRs(rs)
Call CloseDB(conn)
%>
