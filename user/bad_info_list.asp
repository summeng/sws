<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/tt_auto_lock.asp" -->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->

<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/style.css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>不良信息举报</title>
<BODY background="../images/back.gif">
<body>
<div align="center">
<table border="0" cellpadding="0" cellspacing="0" bgcolor="#ffffff" width="980">
<%
dim k,infoid,biaoti
dim ThisPage,Pagesize,Allrecord,Allpage
k=1
infoid=trim(request("infoid"))
biaoti=trim(request("biaoti"))
Call OpenConn
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_Bad_info] where sh=1 and infoid="&infoid&" order by addsj "&DN_OrderType&"" 
rs.open sql,conn,1,1
if rs.eof then
response.write "<p><br><p><table width='980' border='0' align='center' cellpadding='3' cellspacing='1' bgcolor='#000000'><tr><td width='90%' height='25' align='center' bgcolor='#E3EBF9'><li>此信息没有人举报不良记录</td></tr></table>"
response.end
end if
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end if
rs.Pagesize=30
Pagesize=rs.Pagesize
Allrecord=rs.Recordcount
Allpage=rs.Pagecount
if ThisPage<1 then                           
ThisPage=1
end if
rs.move (ThisPage-1)*Pagesize
%>
<tr>

<tr>
<td align="center" >
网上访客对信息[<a href="../show.asp?id=<%=cstr(infoid)%>" target="blank"><%=biaoti%></a>]的举报<br>
（举报内容不一定中肯，仅供参考）
</td>
</tr>
<table width="980" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#000000">
<tr>
<td width="40" height="25" align="center" bgcolor="#E3EBF9">ID</td>
<td width="220" height="25" align="center" bgcolor="#E3EBF9">举报类型</td>
<td width="500" height="25" align="center" bgcolor="#E3EBF9">举报补充内容</td>
<td width="100" height="25" align="center" bgcolor="#E3EBF9">举报者IP</td>
<td width="100" height="25" align="center" bgcolor="#E3EBF9">举报日期</td>
</tr>
</table>
<table width="980" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#000000">
<%do while not rs.eof%>
<tr bgcolor="#ffffff">
<td width="40" height="22" align="center"><%=rs("id")%></td>
<td width="220" height="22" align="center"><%=rs("Bad_infolx")%></td>
<td width="500" height="22" ><%=HTMLEncode(rs("Bad_info"))%></td>
<td width="100" height="22" align="center">
<%
ip=Split(rs("ip"),".")
%><%=Ip(0)%>.<%=Ip(1)%>.<%=Ip(2)%>.*</td>
<td width="100" height="22" align="center"><%=rs("addsj")%></td>
</tr>
<%
k=k+1
rs.movenext
if k>=Pagesize then exit do
loop
Call CloseRs(rs)
Call CloseDB(conn)
%>
</table>
<tr> 
<td height="25" width="800" align="right">
共有&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;条记录
共 <font color="#CC5200"><%=Allpage%></font> 页
现在是第 
<font color="#CC5200"><%=ThisPage%></font> 页
<%
if ThisPage<2 then     
response.write "<font color=""#808080"">首页</font>&nbsp;"
response.write "<font color=""#808080"">上一页</font>&nbsp;"     
else     
response.write "<a href=?page=1&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&id="&id&">首页</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&id="&id&">上一页</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">下一页</font>&nbsp;"
response.write "<font color=""#808080"">尾页</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&id="&id&">下一页</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&id="&id&">尾页</a>&nbsp;"     
end if
%></td>
</tr>



</table>
</div>