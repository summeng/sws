<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->
<!--#include file=usercookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<meta http-equiv="Content-Language" content="zh-cn">
<title>推荐注册的会员列表</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
  <table width="980" height="371" border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" bgcolor="#FFFFFF" style="border-collapse: collapse">
    <tr>
      <td width="214" height="299" valign="top"><div align="center">
    <!--#include file=userleft.asp-->　
        </div></td>
<td width="5">&nbsp;</td>
      <td width="760" height="299" align="center" valign="top"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="100" class="ty1">
  <TR> 
    <TD height="5" colspan="12" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">由会员 <%=username%> 推荐到本站注册的会员列表</FONT></TD>
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
		response.write "<tr><td><div align='center'>没有任何推荐的人</div></td></tr>" 
	else 
	do while not rs.eof%>
	【<a href=../preview.asp?username=<%= rs("username") %> target="_blank"><%=rs("username")%></a>】&nbsp;&nbsp;
    <% 
	rs.movenext 
	loop 
	end if	
	%></td>
    </tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">
<TR>
	<TD>获得本站奖励积分 <%=rs.Recordcount*jf_5%> 分</TD>
</TR>
</TABLE>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">
<TR>
	<TD><font color=#FF0000>您可推荐其它人来本站注册，所推荐的人注册成功您将获得积分<%=jf_5%>分。请将推广代码复制发给您的朋友：</font><br>代码一（放在浏览器地址栏回车）： http://<%=web%>/user_regagree.asp?tjname=<%=tjhy%><br>代码二（适合放在Htm网页、论坛或邮件上）： &lt;a href=http://<%=web%>/user_regagree.asp?tjname=<%=tjhy%> target="_blank">欢迎您到<%=webname%>发布查询供求信息&lt;/a&gt;</TD>
</TR>
</TABLE>
</td>
<%Call CloseRs(rs)
Call CloseDB(conn)%>
