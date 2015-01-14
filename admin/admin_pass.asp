<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../Include/md5.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("01")%>
<script language=javascript src=../Include/mouse_on_title.js></script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<style type="text/css">
<!--
.style1 {
	color: #FFFFFF;
	font-weight: bold;
}
body {
	background-color: #E3EBF9;
}
-->
</style>
<script language="javascript">
<!--
//字节数限制，与下面配合=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}
// --></script>
<script language="javascript" type="text/javascript">
<!--
function CheckForm()
{

        if(document.thisForm.Pass0.value.length<1)
	{
	    alert("原密码不能为空!");
		document.thisForm.Pass0.focus()
	    return false;
	}
	  if ((document.thisForm.Pass1.value.Length()>16) || (document.thisForm.Pass1.value.Length()<8)) //字节数限制，与上面配合
     {
	 alert("新密码长度要求8－16字节，请重新输入！");
	  document.thisForm.Pass1.focus()
	  return false
  }
	    if(document.thisForm.Pass2.value.length<1)
	{
	    alert("确认新密码不能为空!");
		document.thisForm.Pass2.focus()
	    return false;
	}
	   if(document.thisForm.Pass2.value!=document.thisForm.Pass1.value)
        {
            alert("两次输入密码不一致!");
			document.thisForm.Pass2.focus()
            return false;
        }

  
}
//-->
</SCRIPT>
</head>
<BODY>
<BODY background="images/back.gif">

<%dim admin,action,checktext,pass3
admin=request("admin")
action=request("edit")
if action<>"ok" then
%>

	<table width="98%" border="1" 
		style="border-collapse: collapse; border-style: dotted; border-width: 0px"  
		bordercolor="#333333" cellspacing="0" cellpadding="2">
	<form name="thisForm" action=admin_pass.asp method=post  onsubmit="return CheckForm()">
	<tr class=backs><td colspan=2 class=td height=18>管理员密码修改</td></tr>
	<tr><td width=20% align=right height="18">管理员名称</td>
		<td><input type="text" readonly name="admin" size="20" value="<%=admin%>">
  &nbsp;&nbsp; <img src=../images/memo.gif alt="管理员名称不能修改<br>可添加新管理员后删除老管理员"> </td></TR>
	<tr><td width=20% align=right height="18">请您输入旧密码</td>
		<td><input type="password" name="Pass0" size="20"> </td></tr>
	<tr><td width=20% align=right height="18">请您输入新密码</td>
		<td><input type="password" name="Pass1" size="20"> 
  &nbsp;&nbsp; <img src=../images/memo.gif alt="密码长度：8-16位<br>建议使用数字和字母组合"> </td></TR>
	<tr><td width=20% align=right height="18">请您确认新密码</td>
		<td><input type="password" name="Pass2" size="20">
  &nbsp;&nbsp; <img src=../images/memo.gif alt="确认新密码"> </td></TR>
	<tr><td colspan=2><input type="submit" name="Submit" value="确认修改">
		<input type="hidden" name="edit" value="ok"></td></tr>
	</form>
	</table>
<%   
else
	

	if trim(Request("pass0"))="" or trim(Request("pass1"))="" or trim(Request("pass2"))="" then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，填写不完整，请填写旧密码和新密码！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if trim(Request("pass1"))<>trim(Request("pass2")) then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，两次输入的新密码不符，请检查后重新输入！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if len(trim(Request("pass1")))<8 then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，您输入的新密码太短了，要求长度为8-16位，建议使用数字和字母组合！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

 Call OpenConn 
	Set rs = conn.Execute("select * from DNJY_admin where username='"&admin&"'")   
	if rs("password")<>md5(Request("pass0")) then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，旧密码不正确，请检查后重新输入！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from DNJY_admin where username='"&admin&"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		response.write "<script language='javascript'>"
		response.write "alert('操作失败。您未登陆，或者闲置时间过长，或者数据库中无此记录。');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
	else
		rs("password")=md5(Request("pass1"))
		rs.update

	end if
	set rs=nothing
Call CloseDB(conn)
	response.write "<script language='javascript'>"
	response.write "alert('密码更新成功，请牢牢记住您的新密码！！\n\n现在将退出管理中心，请用新密码重新登录！');"
	response.write "location.href='exit.asp';"			
	response.write "</script>"

end if

%>
