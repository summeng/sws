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
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<script language=javascript src=Include/mouse_on_title.js></script>
<style type="text/css">
<!--
.style1 {color: #FF0000}
.style5 {color: #FF0000; font-weight: bold; }
-->
</style>
<script language="javascript">
<!--
function chkinput(){
if (document.msgadd.name.value==0)
	{
	  alert("请填接收会员ID号！")
	  document.msgadd.name.focus()
	  return false
	 }
if (document.msgadd.neirong.value==0)
	{
	  alert("请填消息内容！")
	  document.msgadd.neirong.focus()
	  return false
	 }
if (document.msgadd.validate_codee.value==0)
	{
	  alert("请填验证码！")
	  document.msgadd.validate_codee.focus()
	  return false
	 }
}
// --></script>
</head>

<body background="img/bg1.gif" leftmargin="0" topmargin="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="980" height="426">
    <tr>
      <td width="214" height="354" valign="top"bgcolor="#FFFFFF"><div align="center">
      
 <!--#include file=userleft.asp--></div></td>
 <td width="5">&nbsp;</td>
 <td width="760" height="354" valign="top" bgcolor="#FFFFFF" class="ty1">
 <%Dim rsmess,pages,allPages,page,neirong,riqi,isnew,name,fname,maxlength,rsuser,username2,huifu
if request.cookies("dick")("username")="" then
response.Write "<center>请先登录</center>"
response.End
end if
username=request.cookies("dick")("username")
fname=request.Form("fname")
username2=request.Form("username2")
name=request("name")
id=request.Form("id")
%>
<%If request("action")="msgsave" Then
    if CInt(trim(request.form("validate_codee")))<>CInt(Session("ttpSN")) then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，验证码不对！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
    Session("ttpSN")=""
    response.end
    end if
    if request.form("name")="" then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，用户ID号没填怎能给他发短信?');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end If
    if request.form("name")="ALL"  Or request.form("name")="all" then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，不能用系统专用用户名ALL！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end If
	set rsuser=server.createobject("adodb.recordset")
	rsuser.Open "select username from [DNJY_user] where username='"&request("name")&"'",conn,1,1	
	if rsuser.eof and rsuser.bof then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，没有找到"&request("name")&"会员，请检查后重新提交！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if	
    if request.form("neirong")=""  then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，没有填写要发送的内容！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	rsuser.close
    set rsuser=nothing
	response.end
	end if

	if len(request.form("neirong"))>300  then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，您填写的内容太长了，请检查后重新提交！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	sql="select * from DNJY_Message"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
rs.AddNew
  	rs("name")=name
	neirong=request.form("neirong")
	neirong=replace(neirong,"=======","<font color=darkgray>======")
  	rs("neirong")=HTMLEncode(neirong)
  	rs("riqi")=now()
 	rs("fname")=request.cookies("dick")("username")
rs.Update
Call CloseRs(rs)

	response.write "<script language='javascript'>"
	response.write "alert('操作成功，已向"&request("name")&"发送一条站内消息！');"
	response.write "location.href='messs.asp?"&C_ID&"';"
	response.write "</script>"
	response.end
conn.close
set conn=Nothing
else%>
<form name=msgadd method=post action=?action=msgsave&<%=C_ID%> onsubmit="return chkinput()">
<%if id<>"" then
Set rs = conn.Execute("select * from DNJY_Message where id="&id) 
if not (rs.eof and rs.bof) then
'避免通过回复消息的方式偷看别人的消息
if rs("name")=request.Cookies(ttSavename&"_shop")("username") then
huifu=vbcrlf&vbcrlf&vbcrlf&vbcrlf&"======="&rs("fname")&"于"&rs("riqi")&"发送的消息原文======"&vbcrlf&replace(rs("neirong"),"<font color=darkgray>","")
end if
end if
set rs=nothing
end if
%>
<script LANGUAGE=JavaScript>
  function textLimitCheck(thisArea, maxLength){
    if (thisArea.value.length > maxLength){
      alert(maxLength + ' 个字限制. \r超出的将自动去除.');
      thisArea.value = thisArea.value.substring(0, maxLength);
      thisArea.focus();
    }
    /*回写span的值，当前填写文字的数量*/
    messageCount.innerText = thisArea.value.length;
  }
</script>	
<table width="760" border="0" cellpadding="4" cellspacing="0" bordercolor="#C0C0C0" align=center bordercolordark="#FFFFFF">
<tr><td>发送站内消息：</tr>
	<tr><td width=20%  height="18">发 给 谁&nbsp;
	<input type=text 
	<%
	if request("fname")<>"" then response.write "value='"&fname&"'"
	if request("username2")<>"" then response.write "value='"&username2&"'"
	if request("name")<>"" then response.write "value='"&name&"'"
	%>"
	name="name" size=18 maxlength=16> （请填写本站注册会员的ID）</td></tr>
<tr><td>
  <textarea rows="15" name="neirong" cols="75" style="BORDER: darkgray 1px solid; FONT-SIZE: 8pt; FONT-FAMILY: verdana ; overflow:auto;" onkeyUp="textLimitCheck(this,300);"><%=huifu%></textarea><br>限<%=300%>个字符 已输入 <font color="#CC0000"><span id="messageCount">0</span></font> 个字 
&nbsp;&nbsp;&nbsp;屏蔽htm语言</td></tr>
<tr><td>算式验证：<img src="../tt_getcodee.asp" alt="验证码,看不清楚?请点击刷新验证码" height="10" style="cursor : pointer;" onClick="this.src='../tt_getcodee.asp?t='+(new Date().getTime());" /><input type="text" name="validate_codee" size="4" onkeyup="value=value.replace(/[^\d|^\.]/g,'')"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))">&nbsp;&nbsp; <img src=../images/memo.gif alt="验证码,看不清楚?请点击刷新验证码"></td></tr>
<tr><td height=50 align=center><input type="submit" value="立即发送" name="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" value="重新填写"><input name="send" type="hidden" value="ok"></td></tr>
</form></td>
</table>      
    </tr>
    <tr>
      <td height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
  <%End if%>
</div>
</body>
</html>