<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("04")%>

<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>修改信息</title>

</head>
<body>
<br><table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">

  <TR> 
    <TD height="5" colspan="12" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">请选择修改文章的类别</FONT></TD>
  </TR>

    <%Dim rs2,sql2  
    set rs =server.createobject("adodb.recordset")	
	sql="select * From DNJY_userNewsClass" 
	rs.open sql,conn,1,1 
	if rs.bof and rs.eof then 
		response.write "<tr><td><div align='center'>没有任何分类</div></td></tr>" 
	else 
		do while not rs.eof
		set rs2=server.createobject("adodb.recordset")
		sql2="select * from DNJY_userNews where ClassID="&rs("id")&" and ifshow=0"
		rs2.open sql2,conn,1,1%>
	
	<tr>
      <td width="15%"><div align="center"><a href="UserNews_Modifylist.asp?isse=0&ClassName=<%=rs("ClassName")%>"><%=rs("ClassName")%></a> （栏目ID号：<%=rs("id")%>）（本类有待审核资讯 <%=rs2.recordcount%> 条）</div></td>
    </tr>
    <% 
	rs.movenext 
	loop 
	end If
Call CloseRs(rs2)	
Call CloseRs(rs)
Call CloseDB(conn)
	%>
</table>

</body>

</html>
