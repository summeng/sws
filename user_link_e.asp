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
<%
dim rsurl,sqlurl
id=trim(request("selectedid"))
if not isnumeric(id) or id="" then
response.write "<li>参数错误！"
response.end
end If
Call OpenConn
set rsurl = Server.CreateObject("ADODB.RecordSet")
sqlurl="select * from [DNJY_comlink] where id="&cstr(id) 
rsurl.open sqlurl,conn,1,1
%>
<SCRIPT language=javascript>
<!--
function CheckForm()
{
if(document.thisForm.web.value.length<1)
	{
	    alert("网站名称不能为空!");
	    return false;
	}
	if(document.thisForm.url.value.length<1)
	{
	    alert("网站地址不能为空!");
	    return false;
	}
			if(document.thisForm.webabout.value.length<1)
	{
	    alert("网站简介不能为空!");
	    return false;
	}
	


}

//-->
</SCRIPT>
<body style="font-size: 9pt">

<div align="center">
<table border="0" width="980" id="table1" bgcolor="#FFFFFF" style="border-left: 1px solid #999999; border-right: 1px solid #999999; padding: 1px; border-top-width:1px; border-bottom-width:1px" cellspacing="0" cellpadding="0">
<tr>
<td width="160" height="25" class=td1>
<p align="center">会员管理中心</font>-友情链接管理</td>
</tr>
<tr>
<td bgcolor="#F6F6F6" valign="top" width="160"><!--#include file=userleft.asp--></td>
<td valign="top" width="560" bgcolor="#F6F6F6">

<table border="0" width="600" cellspacing="0" cellpadding="0" id="table2">
<form method="POST" name="thisForm" action="user_link_chk.asp?selectedid=<%=id%>&<%=C_ID%>">
	<p style="line-height: 220%; margin-top: 0; margin-bottom: 0">
	<b>修改网站友情连接资料</b></p><hr color="#C0C0C0" size="1">
	<p style="line-height: 220%; margin-top: 0; margin-bottom: 0">
	网站名称:<input type="text" name="web" size="30" maxlength="20" value="<%=rsurl("web")%>"> *</p>
	<p style="line-height: 220%; margin-top: 0; margin-bottom: 0">
	网站地址:<input type="text" name="url" size="30" value="<%=rsurl("url")%>"> *</p>
	<p style="line-height: 220%; margin-top: 0; margin-bottom: 0">
	网站简介:<input type="text" name="webabout" size="38" maxlength="80" value="<%=rsurl("webabout")%>"> *</p>
	<p style="line-height: 220%; margin-top: 0; margin-bottom: 0">
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<input border="0" onclick="javascript:return CheckForm();" type="submit" value="提交" name="I2">
	</p>
</form>

</body>
<%rsurl.close
set rsurl=Nothing
Call CloseDB(conn)%>
</table>
</td>
<!--#include file=footer.asp-->
</html>