<!--#include file="../conn.asp"-->
<!--#include file=../config.asp-->
<!--#include file=../Include/mail.asp-->
<!--#include file=cookies.asp-->

<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("08")%>
<%Dim Submit,aJMail,mailbody,MailTitle,UseHtml,Target,thisEmail,sssstemp,emails,SplitEmail,CanHTML,i,tomail,url_return,C_M_Chk
Set aJMail=Server.CreateObject("JMail.Message")   
If aJMail Is Nothing Then  
response.write "<script LANGUAGE='javascript'>alert('温馨小提示：服务器不支持JMail.Message组件，请立即安装此组件，此组件用于邮件发送!');history.go(-1);</script>"
Response.End
End If
If Request("mailbody") <> "" and Request("mailbodyhtm") <> "" Then  
response.write "<script LANGUAGE='javascript'>alert('温馨小提示：两种格式邮件内容不能同时填写!');history.go(-1);</script>"
Response.End
End If
If Request("UseHtml") = "text" and Request.form("mailbody")="" Then  
response.write "<script LANGUAGE='javascript'>alert('温馨小提示：选择文本格式邮件请将内容填写在文本格式邮件内容框!');history.go(-1);</script>"
Response.End
End If
If Request("UseHtml") = "HTML" and Request("mailbodyhtm")="" Then  
response.write "<script LANGUAGE='javascript'>alert('温馨小提示：选择HTML格式邮件请将内容填写在HTML格式邮件内容框!');history.go(-1);</script>"
Response.End
End If
If Request("Target")="custom" and Request("usermails") = "" Then  
response.write "<script LANGUAGE='javascript'>alert('温馨小提示：选择自定义邮箱必须填写邮箱!');history.go(-1);</script>"
Response.End
End If

If Request.form("mailbody") <> "" then
mailbody=Request.form("mailbody")
Else
mailbody=Request.form("mailbodyhtm")
End if
server.ScriptTimeout=99999
Submit=request("submit")

if Submit<>"" Then
If Request("MailTitle") = "" Then  
response.write "<script LANGUAGE='javascript'>alert('温馨小提示：邮件标题不能为空!');history.go(-1);</script>"
Response.End
End If
If mailbody = "" Then  
response.write "<script LANGUAGE='javascript'>alert('温馨小提示：邮件内容不能为空!');history.go(-1);</script>"
Response.End
End If

	MailTitle=Request("MailTitle")	
	UseHtml=Request("UseHtml")		
	Target=Request("Target")

	if Target="ALL" then'全部会员		
        emails=AllEmail("")
	elseif Target="ordinary" then'普通会员
		emails=ordinaryEmail("")
	elseif Target="VIP" then'VIP会员
		emails=vipEmail("")
	elseif Target="store" then'店铺会员
		emails=storeEmail("")
	elseif Target="maillist" then'订阅杂志会员
		emails=maillistEmail("")
	elseif Target="custom" then'自定义邮箱
		emails=Request("usermails")
	end if

	if Target<>"custom" then'如果不是自定义按|数组
	SplitEmail=split(emails,"|")
	else
	SplitEmail=split(emails,chr(13))'自定义按空格数组
	end if
	if UseHtml="" then
		CanHTML=false
	else
		CanHTML=true
	end If	
    C_M_Chk=True'类型：布尔值。作用：Smtp服务器是否需要身份验证 ,True需要验证，False不需要验证。
	for i=0 to ubound(SplitEmail)
		tomail=trim(trim(SplitEmail(i)))
		if instr(tomail,"@")>0 And ChkMail(tomail)=True then			
			Call DNJYEmail(tomail,webname,MailTitle,mailbody,CanHTML,C_M_Chk)'调用邮件函数发送邮件
			response.write "."
		end If		
	Next
	response.write "<script LANGUAGE='javascript'>alert('发送成功！一共"&ubound(SplitEmail)+1&"封邮件!');history.go(-1);</script>"	
end if

'获取邮箱地址＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function AllEmail(userid)'获取全部会员邮箱
    set rs=server.createobject("adodb.recordset")
	sql="select email from DNJY_user where useryz=1"
	rs.open sql,conn,1,1	
	if not rs.eof Then
		do while not rs.eof
			thisEmail=rs(0)
			if instr(sssstemp,thisEmail)=0 then
			sssstemp=sssstemp&thisEmail&"|"
			end if
		rs.movenext
		loop
	end if
	rs.close
	AllEmail=sssstemp
end Function

function ordinaryEmail(userid)'获取普通会员邮箱
    set rs=server.createobject("adodb.recordset")
	sql="select email from DNJY_user where useryz=1 and vip=0"
	rs.open sql,conn,1,1	
	if not rs.eof Then
		do while not rs.eof
			thisEmail=rs(0)
			if instr(sssstemp,thisEmail)=0 then
			sssstemp=sssstemp&thisEmail&"|"
			end if
		rs.movenext
		loop
	end if
	rs.close
	ordinaryEmail=sssstemp
end Function

function vipEmail(userid)'获取VIP会员邮箱
    set rs=server.createobject("adodb.recordset")
	sql="select email from DNJY_user where useryz=1 and vip=1"
	rs.open sql,conn,1,1	
	if not rs.eof Then
		do while not rs.eof
			thisEmail=rs(0)
			if instr(sssstemp,thisEmail)=0 then
			sssstemp=sssstemp&thisEmail&"|"
			end if
		rs.movenext
		loop
	end if
	rs.close
	vipEmail=sssstemp
end Function

function storeEmail(userid)'获取店铺会员邮箱
    set rs=server.createobject("adodb.recordset")
	sql="select email from DNJY_user where useryz=1 and sdname<>''"
	rs.open sql,conn,1,1	
	if not rs.eof Then
		do while not rs.eof
			thisEmail=rs(0)
			if instr(sssstemp,thisEmail)=0 then
			sssstemp=sssstemp&thisEmail&"|"
			end if
		rs.movenext
		loop
	end if
	rs.close
	storeEmail=sssstemp
end Function

function maillistEmail(userid)'获取订阅杂志会员邮箱
    set rs=server.createobject("adodb.recordset")
	sql="select email from DNJY_user where useryz=1 and maillist=1"
	rs.open sql,conn,1,1	
	if not rs.eof Then
		do while not rs.eof
			thisEmail=rs(0)
			if instr(sssstemp,thisEmail)=0 then
			sssstemp=sssstemp&thisEmail&"|"
			end if
		rs.movenext
		loop
	end if
	rs.close
	maillistEmail=sssstemp
end Function

'获取邮箱地址END＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝%>
<html>
<head>
<title>邮件群发</title>
<meta http-equiv="目录类型" content="文本/html; 字符集=gb2312">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<!--kindeditor编辑器-->
		<link rel="stylesheet" href="../KindEditor/themes/default/default.css" />
		<script charset="utf-8" src="../KindEditor/kindeditor-min.js"></script>
		<script charset="utf-8" src="../KindEditor/lang/zh_CN.js"></script>
		<script>
			var editor;			
			KindEditor.ready(function(K) {
				editor = K.create('textarea[name="mailbodyhtm"]', {
					resizeType : 2,//0:设置kindeditor编辑器大小不可改变; 1:只能该变高度; 2:可以改变高度和宽度
					afterBlur: function(){this.sync();},//失去焦点执行this获得值
					allowPreviewEmoticons : false,
					allowImageUpload : false,					
					items : [
						'fontname', 'fontsize', '|', 'forecolor', 'hilitecolor', 'bold', 'italic', 'underline',
						'removeformat', '|', 'justifyleft', 'justifycenter', 'justifyright', 'insertorderedlist',
						'insertunorderedlist', '|', 'emoticons', 'image', 'link']
				});
			});
		</script>
<!--kindeditor编辑器END-->
<script language="javascript">
<!--
//验证邮箱正确性
function checkemail(email){
var str=email;
//在JavaScript中，正则表达式只能使用"/"开头和结束，不能使用双引号
var Expression=/\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*/; 
var objExp=new RegExp(Expression);
if(objExp.test(str)==true){
return true;
}else{
return false;
}
} 

//说明字节数限制，与下面配合=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}

function formcheck(){
    if((!checkemail(form1.usermails.value))&&(document.form1.usermails.value!=""))
	{
    alert("您输入Email地址不正确，请重新输入!");
    document.form1.usermails.focus();
    return false;
    }
if (document.form1.MailTitle.value=="" )
	{
	  alert("请填邮件主题")
	  document.form1.MailTitle.focus()
	  return false
	 }
if (document.sendmail.mailbody.value.length>5000)  //字节数限制，与上面配合。
     {
	 alert("邮件内容长度要求<%=5000%>字节(<%=2500%>汉字)内，请重新输入！");
	  return false
  }  
}
// --></script>
</head>
<BODY>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
.style1 {
	color: #FFFFFF;
	font-weight: bold;
	font-size: 12px;
}
.style2 {font-size: 12px}
-->
</style>
<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>
<tr><td colspan=4 bgcolor="#BDBEDE" height=28 align=center><B>群发邮件</b></td></tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="0" class="border">
 <form name="form1" method="post" action="" onSubmit="return formcheck()">
 <tr> <td width="19%" align="center" colspan="2" class="tdbg"><font color="#ff0000"> 注意：由于很多邮局、尤其是免费邮箱群发邮件都受监测，可能因过密或内容相同而被暂时或永久禁止发信邮箱的发信功能，如果需经常群发邮件的建议购买不受限制的企业邮局！</font></td></tr> 
<tr> <td width="19%" align="right" class="tdbg"><b>目&nbsp;&nbsp;&nbsp;&nbsp;标：</b> </td><td width="81%" class="tdbg"> <input name="Target" type="radio" value="ALL" checked> 
全部会员 <input type="radio" name="Target" value="ordinary"> 普通会员 <input type="radio" name="Target" value="VIP"> VIP会员 <input type="radio" name="Target" value="store">店铺会员 <input type="radio" name="Target" value="maillist"> 订阅杂志会员 <input type="radio" name="Target" value="custom"> 
根据自定义收件人邮箱</td></tr> 
<tr> <td width="19%" align="right" class="tdbg"><b>自定义收件人邮箱：</b></td><td width="100%" class="tdbg"><textarea name="usermails" cols="84" rows="9" wrap="VIRTUAL" id="usermails"></textarea>
    （<font color="#ff0000">邮箱每行一个，回车添加，规则不对无法发送！</font>）<p></td></tr>

<tr> 
<td width="19%" align="right" class="tdbg"><b>邮件格式：</b></td><td width="81%" class="tdbg"><input type="radio" name="UseHtml" value="text" checked  style="border:1px black solid">                                                                   
      文本邮件&nbsp; <input type="radio" name="UseHtml" value="HTML" style="border:1px black solid">                                                                          
        HTML邮件&nbsp;<p></td></tr> 
<tr> <td width="19%" align="right" class="tdbg"><b>邮件标题：</b></td><td width="81%" class="tdbg"> 
<input name="MailTitle" type="text" id="MailTitle" size="60"> <font color="#ff0000"> *</font><p></td></tr> 
<tr> <td width="19%" align="right" class="tdbg"></td><td width="81%" class="tdbg"><font color="#ff0000"> 下面两个内容框要根据选择的邮件格式对应填写，不能同时填写！</font> </td></tr> 
<tr> <td width="19%" align="right" valign="top" class="tdbg"><b>文本格式邮件内容：</b></td><td width="81%" class="tdbg"> 
<textarea name="mailbody" cols="84" rows="15" wrap="VIRTUAL" id="mailbody"></textarea> 
 <font color="#ff0000"> *</font></td></tr>
<tr> <td width="19%" align="right" valign="top" class="tdbg"><b>HTML格式邮件内容：</b></td><td width="81%" class="tdbg"> 
<textarea id="editor" name="mailbodyhtm"  style="width:550px;height:280px;visibility:visible"></textarea>
 <font color="#ff0000"> *</font></td></tr>
<tr> <td width="19%" height="36" align="right" class="tdbg">&nbsp;</td><td width="81%" height="36" class="tdbg"> 
<input type="submit" name="Submit" value="确定并发送邮件  "> <input type="reset" name="Submit2" value="  全部重写  "> 
</td></tr>

<tr>
  <td align="right" class="tdbg">&nbsp;</td>
  <td class="tdbg">&nbsp;</td>
</tr>
<tr>
  <td align="right" class="tdbg">&nbsp;</td>
  <td class="tdbg">&nbsp;</td>
</tr>
</form></table>
</body>
</html>
<!--#include virtual="/Include/bottom_superadmin.asp" -->
