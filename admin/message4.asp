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
<title>发件箱管理</title>
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<BODY>
</head>
<script language=javascript src=../Include/mouse_on_title.js></script>

<%Dim sql2,url,id,str2,i
Call OpenConn
	if request("delid")<>"" then	
	sql2= "delete from DNJY_Message where id="&request("delid")
	url="Message1.asp"
	end if

if request("delid3")<>"" Then
id=trim(request("delid3"))
str2=split(id,",")
for i=0 to ubound(str2)'循环开始
set rs=server.createobject("adodb.recordset")
sql2="delete from [DNJY_Message] where id="&cstr(str2(i))
rs.open sql2,conn,1,3
set rs=Nothing
Next'循环结束
	url="Message3.asp"
	end if

	if request("delid5")<>"" then	
	sql2= "delete from DNJY_Message where id="&request("delid5")
	url="Message5.asp"
	end if

	if request("action")="all" then
	sql2= "delete from DNJY_Message"
	url="Message4.asp"
	end if

	if request("action")="w" then      
	sql2="delete from DNJY_Message where riqi<"&DN_Now&"-7"
	url="Message4.asp"
	end if

	if request("action")="m" then
	sql2="delete from DNJY_Message where riqi<"&DN_Now&"-30"
	url="Message4.asp"
	end if
	
	if request("action")="isnew" then
	sql2="delete from DNJY_Message where isnew='1'"
	url="Message4.asp"
	end If
	
	if request("action")="hydel" then
	sql2="delete from DNJY_Message where del='1'"
	url="Message4.asp"
	end if

if sql2<>"" then
	conn.execute (sql2)
	response.write "<script language='javascript'>"
	response.write "alert('所选站内消息删除成功！');"
	response.write "location.href='"&url&"';"
	response.write "</script>"
	response.end
else
%>


	<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#333333" cellspacing="0" cellpadding="2">
	<form action=Message4.asp method=post name="hand">
	<tr class=backs><td colspan=2 class=td height=18>删除站内消息</td></tr>

	<tr><td width=20% align=right height="18">删除：&nbsp;</td><td>
	<input type=radio value="w" name="action" checked>一周前 
	<input type=radio value="m" name="action">一月前 
	<input type=radio value="isnew" name="action">所有已阅读消息
	<input type=radio value="hydel" name="action">会员已删除消息 
	<input type=radio value="all" name="action">所有消息 
	</td></tr>

	<tr><td colspan=2>
	<INPUT name=del TYPE="submit" value="立即删除" onclick="{if(confirm('此操作不可恢复，您确认要删除所选信息吗？\n\n单击确定继续，单击取消返回。')){this.document.hand.submit();return true;}return false;}"></td></tr>
	</form>
	</table>


</body>
</html>
<%end If
Call CloseDB(conn)%>
