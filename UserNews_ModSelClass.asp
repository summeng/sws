<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=usercookies.asp-->
<!--#include file=top.asp-->

<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<html>
<head>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�޸���Ϣ</title>

</head>
<body>
  <table width="980" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
    <tr>
      <td width="214" height="356" valign="top" bgcolor="#FFFFFF"><div align="center"><!--#include file=userleft.asp--></div></td>
	  <td width="5">&nbsp;</td>
      <td width="860" align="center" valign="top" bgcolor="#FFFFFF">
	<!---�ϲ�ͨ����濪ʼ---->
  <table align="center" bgcolor="#ffffff">
  <tr>  
   <td ><script language=javascript src="<%=adjs_path("ads/js","fygg",c1&"_"&c2&"_"&c3)%>"></script></td>
  </tr>
  </table>
  <!---�ϲ�ͨ��������----> 
<br><table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">

  <TR> 
    <TD height="5" colspan="12" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">��ѡ���޸����µ����</FONT></TD>
  </TR>

    <%Call OpenConn  
    set rs =server.createobject("adodb.recordset")	
	sql="select * From DNJY_userNewsClass" 
	rs.open sql,conn,1,1 
	if rs.bof and rs.eof then 
		response.write "<tr><td><div align='center'>û���κη���</div></td></tr>" 
	else 
		do while not rs.eof 
	%>
	
	<tr>
      <td width="15%"><div align="center"><a href="UserNews_Modifylist.asp?isse=0&ClassName=<%=rs("ClassName")%>"><%=rs("ClassName")%></a></div></td>
    </tr>
    <% 
	rs.movenext 
	loop 
	end if 
	rs.close 
Call CloseDB(conn)
	%>
	</table>
	</td>
	  </tr>
</table>
</body>
</html>
