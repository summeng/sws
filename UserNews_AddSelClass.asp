<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=usercookies.asp-->
<!--#include file=top.asp-->

<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<html>
<head>
<link rel="stylesheet" type="text/css" href="css/css.css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>添加文章</title>
</head>
<body>
<br>
  <table width="980" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
    <tr>
      <td width="214" height="356" valign="top" bgcolor="#FFFFFF"><div align="center"><!--#include file=userleft.asp--></div></td>
	  <td width="5">&nbsp;</td>
      <td width="860" align="center" valign="top" bgcolor="#FFFFFF">
	<!---上部通栏广告开始---->
  <table align="center" bgcolor="#ffffff">
  <tr>  
   <td ><script language=javascript src="<%=adjs_path("ads/js","fygg",c1&"_"&c2&"_"&c3)%>"></script></td>
  </tr>
  </table>
  <!---上部通栏广告结束---->
<%Dim sdkg
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_user] where username='"&username&"'" 
rs.open sql,conn,1,1
if rs.eof then
response.write "<br>"
response.write "<li>参数错误！"
response.end
end if
sdkg=rs("sdkg")
vip=rs("vip")
rs.close
If sdkg=0 then%>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">
  <TR> 
    <TD height="5" colspan="12" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">您没有开店铺,或店铺未经审核,不能发布文章!</FONT></TD>
  </TR>
 </table>
 <% response.End
 End If%>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">
  <TR> 
    <TD height="5" colspan="12" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">请选择发布文章的类别</FONT></TD>
  </TR>
    <%  
    set rs =server.createobject("adodb.recordset")	
	sql="select * From DNJY_userNewsClass" 
	rs.open sql,conn,1,1 
	if rs.bof and rs.eof then 
		response.write "<tr><td><div align='center'>没有任何分类</div></td></tr>" 
	else 
		do while not rs.eof 
	%>	
	<tr>
      <td width="15%"><div align="center"><a href="UserNews_Add.asp?ClassName=<%=rs("ClassName")%>"><%=rs("ClassName")%></a></div></td>
    </tr>
    <% 
	rs.movenext 
	loop 
	end if 
	rs.close
	set rs = nothing
	conn.close
	set conn=nothing
	%>
	</table>(发布文章的条件是必须开有店铺,且店铺经审核通过。发布的文章在自己的店铺显示，并可共享到总网站。)
	</td>
	  </tr>
</table>
</body>
</html>
