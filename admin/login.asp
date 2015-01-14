<!--#include file="../config.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%'禁止网页缓存
Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>网站管理平台</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK media=screen href="images/login.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2900.3059" name=GENERATOR></HEAD>
<BODY>
<script language="javascript">
<!--
function myform_onsubmit() {
if (document.myform.username.value=="" )
	{
	  alert("请填用户名！")
	  document.myform.username.focus()
	  return false
	 }
if (document.myform.password.value=="" )
	{
	  alert("请填密码！")
	  document.myform.password.focus()
	  return false
	 }
if (document.myform.admin_codee.value=="" )
	{
	  alert("请填验证码！")
	  document.myform.admin_codee.focus()
  return false
  }
  
}
// --></script>
<TABLE cellSpacing=0 cellPadding=0 width=969 align=center border=0>
 <form name="myform" action="login_save.asp" method="post" target="_parent" onSubmit="return myform_onsubmit()">	
 <TBODY>
  <TR>
    <TD><IMG height=177 alt="" src="images/login_01.gif" width=241></TD>
    <TD><IMG height=177 alt="" src="images/login_02.gif" width=220></TD>
    <TD><IMG height=177 alt="" src="images/login_03.gif" width=216></TD>
    <TD><IMG height=177 alt="" src="images/login_04.gif" width=292></TD></TR>
  <TR>
    <TD colspan="4" vAlign=top><img src="images/login_6-10.jpg" width="969" height="180"></TD>
    </TR>
  <TR>
    <TD><IMG height=190 alt="" src="images/login_10.gif" width=241></TD>
    <TD class=login colSpan=2>
      <P><LABEL>用户帐号</LABEL><INPUT id="admin" name="username">
      </P>
      <P><LABEL>登录密码</LABEL><INPUT id="password" type="password" name="password"></P>
      <br> &nbsp;&nbsp;<LABEL>验证编码</LABEL>&nbsp;&nbsp;   <img src="../tt_getcodee.asp" alt="验证码,看不清楚?请点击刷新验证码" height="10" style="cursor : pointer;" onClick="this.src='../tt_getcodee.asp?t='+(new Date().getTime());" /><input name="admin_codee" id="admin_codee" type="text" tabindex="3" value="" size="17" maxlength="4" /></P>
      <DIV><INPUT class=no_input type=image src="images/login.gif"> 
        <a href="javascript:window.close()"><img src="images/quxiao.gif" border="0"></a> </DIV></TD>
    <TD><IMG height=190 alt="" src="images/login_12.gif" width=292></TD></TR>
  <TR>
    <TD><IMG height=123 alt="" src="images/login_13.gif" width=241></TD>
    <TD style="FONT-WEIGHT: bold; COLOR: #095816" vAlign=bottom align=middle background=images/login_14.gif colSpan=2 height=123>技术支持：<a href="http://www.ip126.com" target=_blank>天天网络科技</a></TD>
    <TD><IMG height=123 alt="" src="images/login_15.gif" width=292></TD></TR></TBODY></form></TABLE>
</BODY>
</HTML>
<%Response.Cookies("status")="onoroff"%>