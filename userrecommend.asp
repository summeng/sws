<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->
<!--#include file=usercookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<meta http-equiv="Content-Language" content="zh-cn">
<title>�Ƽ�ע��Ļ�Ա�б�</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
  <table width="980" height="371" border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" bgcolor="#FFFFFF" style="border-collapse: collapse">
    <tr>
      <td width="214" height="299" valign="top"><div align="center">
    <!--#include file=userleft.asp-->��
        </div></td>
<td width="5">&nbsp;</td>
      <td width="760" height="299" align="center" valign="top"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="100" class="ty1">
  <TR> 
    <TD height="5" colspan="12" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">�ɻ�Ա <%=username%> �Ƽ�����վע��Ļ�Ա�б�</FONT></TD>
  </TR>

	<tr>
      <td width="15%">
    <%username=request.cookies("dick")("username")
    Dim tjhy
    tjhy=username
	Call OpenConn  
    set rs =server.createobject("adodb.recordset")	
	sql="select username From DNJY_user where tjname='"&username&"'" 
	rs.open sql,conn,1,1 
	if rs.bof and rs.eof then 
		response.write "<tr><td><div align='center'>û���κ��Ƽ�����</div></td></tr>" 
	else 
	do while not rs.eof%>
	��<a href=../preview.asp?username=<%= rs("username") %> target="_blank"><%=rs("username")%></a>��&nbsp;&nbsp;
    <% 
	rs.movenext 
	loop 
	end if	
	%></td>
    </tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">
<TR>
	<TD>��ñ�վ�������� <%=rs.Recordcount*jf_5%> ��</TD>
</TR>
</TABLE>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">
<TR>
	<TD><font color=#FF0000>�����Ƽ�����������վע�ᣬ���Ƽ�����ע��ɹ�������û���<%=jf_5%>�֡��뽫�ƹ���븴�Ʒ����������ѣ�</font><br>����һ�������������ַ���س����� http://<%=web%>/user_regagree.asp?tjname=<%=tjhy%><br>��������ʺϷ���Htm��ҳ����̳���ʼ��ϣ��� &lt;a href=http://<%=web%>/user_regagree.asp?tjname=<%=tjhy%> target="_blank">��ӭ����<%=webname%>������ѯ������Ϣ&lt;/a&gt;</TD>
</TR>
</TABLE>
</td>
<%Call CloseRs(rs)
Call CloseDB(conn)%>
