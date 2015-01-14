<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Call OpenConn
dim username
username=trim(request("username"))
if username="" then
response.write "<li>参数错误！"
response.end
end if
%>
<meta http-equiv="Content-Language" content="zh-cn">
<title>推荐注册的会员列表</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<body topmargin="0" leftmargin="0">
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">

  <TR> 
    <TD height="5" colspan="12" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">由会员 <%=username%> 推荐<br>到本站注册的会员列表</FONT></TD>
  </TR>

    <%  
    set rs =server.createobject("adodb.recordset")	
	sql="select username From DNJY_user where tjname='"&username&"'" 
	rs.open sql,conn,1,1 
	if rs.bof and rs.eof then 
		response.write "<tr><td><div align='center'>没有任何推荐的人</div></td></tr>" 
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
	<TD>获得本站奖励积分 <%=rs.Recordcount*jf_5%> 分</TD>
</TR>
</TABLE>
<%Call CloseRs(rs)
Call CloseDB(conn)%>
