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
dim username
username=trim(request("username"))
if username="" then
response.write "<li>��������"
response.end
end if
%>
<meta http-equiv="Content-Language" content="zh-cn">
<title>�Ƽ�ע��Ļ�Ա�б�</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<body topmargin="0" leftmargin="0">
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">

  <TR> 
    <TD height="5" colspan="12" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">�ɻ�Ա <%=username%> �Ƽ�<br>����վע��Ļ�Ա�б�</FONT></TD>
  </TR>

    <%  
    set rs =server.createobject("adodb.recordset")	
	sql="select username From DNJY_user where tjname='"&username&"'" 
	rs.open sql,conn,1,1 
	if rs.bof and rs.eof then 
		response.write "<tr><td><div align='center'>û���κ��Ƽ�����</div></td></tr>" 
	else 
	do while not rs.eof%>	
	<tr>
      <td width="15%"><div align="center"><a href=../preview.asp?username=<%= rs("username") %> target="_blank"><%=rs("username")%></a></div></td>
    </tr>
    <% 
	rs.movenext 
	loop 
	end if	
	%>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">
<TR>
	<TD>��ñ�վ�������� <%=rs.Recordcount*jf_5%> ��</TD>
</TR>
</TABLE>
<%Call CloseRs(rs)
Call CloseDB(conn)%>
