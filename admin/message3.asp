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
<%call checkmanage("09")%>
<html>
<head>
<title>站内短信管理</title>
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<BODY>
</head>
<script language=javascript src=../Include/mouse_on_title.js></script>
<script language="JavaScript">
<!--
function CheckAll(form)
{
for (var i=0;i<form.elements.length;i++)
{
var e = form.elements[i];
if (e.name != 'chkall')
e.checked = form.chkall.checked;
}
}

//-->
</script>
	<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#333333" cellspacing="0" cellpadding="2">
	<tr class=backs><td class=td height=18>站内短信</td></tr>
<%dim pages,allPages,page,neirong,riqi,isnew,del,fname,name,id,temp
Call OpenConn
set rs=Server.CreateObject("ADODB.recordset")
sql="select * from DNJY_Message order by riqi "&DN_OrderType&""
rs.open sql,conn,1,3
if rs.eof and rs.bof then 
response.write "<tr><td>没有站内短信消息。</td></tr>"
else
response.write "<marquee><font color=red>及时删除无用短信，重要资料请及时转移保存，确保安全！</font></marquee>"
pages = 10			'每页记录数
rs.pageSize = pages
allPages = rs.pageCount		'计算一共能分多少页
page = Request("page")
'if语句属于基本的排错处理
if isEmpty(page) or clng(page) < 1 then
page = 1
elseif clng(page) >= allPages then
page = allPages 
end if
rs.AbsolutePage = page%>
<form name="form1" method="post" action="Message4.asp">
<%do while not rs.eof and pages>0
neirong=rs("neirong")	'内容
riqi=rs("riqi")		'日期
isnew=rs("isnew")	'是否阅读
del=rs("del")		'是否删除
fname=rs("fname")   '发件人
name=rs("name")		'收件人
id=rs("id")		'编号
if isnew="0" then
	temp="收件人未阅读"
elseif isnew="1" then
	if del="0" then 
	temp="<font color=red>收件人已阅读</font>"
	elseif del="1" then
	temp="<font color=red>收件人已删除</font>"
	end if	
end if
if name="ALL" then temp="公共短信"
if pages<10 then response.write "<tr><td height=15></td></tr>"
%>


<tr><td onMouseOver="bgColor='#FFFFFF'" onMouseOut="bgColor='#e7e7e7'">
选择<input name="delid3" type="checkbox" id="delid3" value="<%=rs("id")%>">
&nbsp;&nbsp;
<span alt=彻底删除此消息 style="cursor:hand" onclick="{if(confirm('该操作不可恢复！\n\n确实删除这条消息吗？ ')){location.href='Message4.asp?delid3=<%=id%>';}}"><img src=../images/m_delete.gif border=0 align=center></span><span alt=转发给其它会员 style="cursor:hand" onclick="{location.href='Message2.asp?back=kkk3&fw=yes&id=<%=id%>';}"><img src=../images/m_fw.gif border=0 align=center></span>&nbsp;&nbsp;
发件人：<a href="Message2.asp?back=kkk1&id=<%=id%>"><font color=red><%=fname%></font></a>&nbsp;&nbsp;
收件人：<a href="Message2.asp?back=kkk1&name=<%=name%>"><font color=red><%=name%></font></a> &nbsp; &nbsp; &nbsp; 发送时间：<%=riqi%> &nbsp; &nbsp; &nbsp; <b><%=temp%></b>
<hr color=#e7e7e7><BR>
<%=replace(neirong,vbCRLF,"<BR>")%><br>&nbsp;
</td></tr>


<%
pages = pages - 1
rs.movenext
if rs.eof then exit do
loop
response.write "</table>"%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
<tr>
<td height="20" colspan="15" align="center" bgcolor="#F3F3F3">	
  <p align="center">
    <input type="checkbox" name="chkall" onClick="CheckAll(this.form)">
    全部选中 
    <input type="submit" name="Submit" value="删除所选消息" onClick="{if(confirm('该操作不可恢复！\n\n确实删除选定的消息？')){return true;}return false;}">
  </p>
 </td> 
</tr>
</form>
</table>
<%
call listpages()
end if
Call CloseRs(rs)%>

<%
'分页
sub listPages() 
'if allpages <= 1 then exit sub 
if page = 1 then
response.write "<font color=darkgray>首页 前页</font>"
else
response.write "<a href="&request.ServerVariables("script_name")&"?page=1>首页</a> <a href="&request.ServerVariables("script_name")&"?page="&page-1&">前页</a>"
end if
if page = allpages then
response.write "<font color=darkgray> 下页 末页</font>"
else
response.write " <a href="&request.ServerVariables("script_name")&"?page="&page+1&">下页</a> <a href="&request.ServerVariables("script_name")&"?page="&allpages&">末页</a>"
end if
response.write " 第"&page&"页 共"&allpages&"页"
end sub
Call CloseDB(conn)%>


</td></tr></table>
</body>
</html>

