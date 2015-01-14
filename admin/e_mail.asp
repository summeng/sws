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
<html>
<head>
<title>发送电子杂志</title>
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
    if((!checkemail(sendmail.tomail.value))&&(document.sendmail.tomail.value!=""))
	{
    alert("您输入Email地址不正确，请重新输入!");
    document.sendmail.tomail.focus();
    return false;
    }
if (document.sendmail.mailsubject.value=="" )
	{
	  alert("请填邮件主题")
	  document.sendmail.mailsubject.focus()
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
<center>
<%Dim aJMail   
Set aJMail=Server.CreateObject("JMail.Message")   
If aJMail Is Nothing Then  
response.write "<script LANGUAGE='javascript'>alert('温馨小提示：服务器不支持JMail.Message组件，请立即安装此组件，此组件用于邮件发送!');history.go(-1);</script>"
Response.End
End If
If Request.form("tomail") <> "" and Request("qunfa") <> "" Then  
response.write "<script LANGUAGE='javascript'>alert('温馨小提示：填写收件人邮箱后不能同时选择群发!');history.go(-1);</script>"
Response.End
End If
If Request.form("mailbody") <> "" and Request("mailbodyhtm") <> "" Then  
response.write "<script LANGUAGE='javascript'>alert('温馨小提示：两种格式邮件内容不能同时填写!');history.go(-1);</script>"
Response.End
End If
If Request("radiobutton") = "text" and Request("mailbody")="" Then  
response.write "<script LANGUAGE='javascript'>alert('温馨小提示：选择文本格式邮件请将内容填写在文本格式邮件内容框!');history.go(-1);</script>"
Response.End
End If
If Request("radiobutton") = "html" and Request("mailbodyhtm")="" Then  
response.write "<script LANGUAGE='javascript'>alert('温馨小提示：选择HTML格式邮件请将内容填写在HTML格式邮件内容框!');history.go(-1);</script>"
Response.End
End If
If Request.form("mailbody") <> "" then
mailbody=Request.form("mailbody")
Else
mailbody=Request.form("mailbodyhtm")
End if
dim userid,rs1,tomail,mailsubject,mailbody,host,dengji,ErrStr,jmail,sitename,email,i

if Request.form("action")="send" then
call mailto()
Else
call mailedit()
End if
sub mailto()
If ChkMail(Request.form("tomail")) =False Then
response.Write "<script language=javascript>alert('收信人邮箱格式错误!');history.go(-1);</script>"
response.end
End If
if mailbody <> "" and Request.form("mailsubject") <> "" then 
userid=request.querystring("id")
set rs1=server.createobject("adodb.recordset")
rs1.open "select email from DNJY_user where maillist=1 and email<>'' " ,conn,1,1	
	email = Request.form("tomail") 
	mailsubject = Request.form("mailsubject") '信件主题
	mailbody = mailbody '信件内容
	host = ""
	dengji = Request.form("dengji")  '邮件等级
	ErrStr = "ture"	

if Request.form("qunfa") = "" then'决定是否群发
'jmail邮件组件发送邮件开始，下面的值要先取得
Set jmail = Server.CreateObject("JMail.Message") 
jmail.silent = true '屏蔽例外错误，返回FALSE跟TRUE两值
jmail.Logging = true '启用邮件日志
jmail.Charset = "gb2312"  '邮件的文字编码为国标
If Request("radiobutton") = "html" Then
jmail.ContentType = "text/html"    '邮件的格式为HTML格式
end if
jmail.MailServerUserName = mailname '输入smtp服务器验证登陆名 （邮局中任何一个用户的Email地址） 
jmail.MailServerPassword = mailpass '输入smtp服务器验证密码 （用户Email帐号对应的密码） 
jmail.From = mailform '发信人信箱，因服务器反垃圾邮件，此信箱必须与发信服务器同域用户的信箱！ 
jmail.FromName = webname '发件人姓名 
jmail.AddRecipient email '收件人Email 
jmail.Subject = mailsubject '邮件的标题 
jmail.Body = mailbody '邮件的内容
JMail.Priority = dengji      '邮件的紧急程序 
jmail.Send (mailsmtp) '执行邮件发送smtp服务器地址（企业邮局地址）
ErrStr =jmail.errormessage
jmail.close 
set jmail = nothing '关闭对象
'邮件发送结束


else'群发
for i=1 to rs1.recordcount
email = rs1("email")  '客户信箱
If ChkMail(email) =true Then'判断邮箱是否合法
'jmail邮件组件发送邮件开始，下面的值要先取得
Set jmail = Server.CreateObject("JMail.Message") 
jmail.silent = true '屏蔽例外错误，返回FALSE跟TRUE两值
jmail.Charset = "gb2312"  '邮件的文字编码为国标
If Request("radiobutton") = "html" Then
jmail.ContentType = "text/html"    '邮件的格式为HTML格式
end if
jmail.MailServerUserName = mailname '输入smtp服务器验证登陆名 （邮局中任何一个用户的Email地址） 
jmail.MailServerPassword = mailpass '输入smtp服务器验证密码 （用户Email帐号对应的密码） 
jmail.From = mailform '发信人信箱，因服务器反垃圾邮件，此信箱必须与发信服务器同域用户的信箱！ 
jmail.FromName = webname '发件人姓名 
jmail.AddRecipient email '收件人Email 
jmail.Subject = mailsubject '邮件的标题 
jmail.Body = mailbody '邮件的内容
JMail.Priority = dengji      '邮件的紧急程序 
jmail.Send (mailsmtp) '执行邮件发送smtp服务器地址（企业邮局地址）
ErrStr =jmail.errormessage
jmail.close 
set jmail = nothing '关闭对象
'邮件发送结束
end if'判断邮箱是否合法
		rs1.movenext
		next
		rs1.close
	end if	'群发判断结束
if ErrStr = "" then
	Response.Write "恭喜！邮件发送成功！" 
'＝＝＝＝＝同时写邮件日志开始＝＝＝＝
Dim rslog
set rslog=server.createobject("adodb.recordset")
sql="select * from DNJY_log"
rslog.open sql,conn,1,3
rslog.addnew
rslog("username")=request.cookies("administrator")
rslog("outbox")=mailform
rslog("inbox")=email
rslog("title")=mailsubject
rslog("content")=mailbody
rslog("Sendtime")=now()
rslog("lock")=0
rslog.update
rslog.close
set rslog=nothing
'＝＝＝＝同时写邮件日志END＝＝＝＝		
else
	
	Response.Write "很遗憾！邮件未能发送！<br>请检查邮件系统各项设置、是否有订阅杂志会员."
end if	'发送结果判断结束     
Response.end
else		     
response.write "<script language='javascript'>"
response.write "alert('出错了，邮件主题、邮件内容必须完整填写。');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end                                                                          
end if  '是否发信判断结束                                                                         
Call CloseDB(conn)
end Sub
sub mailedit()%> 
<table border="1" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%" cellpadding="0">
  <tr bgcolor="#BDBEDE">
    <td width="100%" height="28" colspan="8" align="center"><b><span class="style2">发送电子杂志</span></b></td>
  </tr>
</table>
  <form name="sendmail" action="e_mail.asp" method="post" onSubmit="return formcheck()">
 <table border="1" width="95%" height="161"  background="images/greystrip.gif"  bordercolordark="#FFFFFF"  cellspacing="0">
      <tr>
    <td width="100%" colspan="2" height="20">
      <p align="center" ><font color="#ff0000"> 注意：由于很多邮局、尤其是免费邮箱群发邮件都受监测，可能因过密或内容相同而被暂时或永久禁止发信功能，如果需经常群发邮件的建议购买不受限制的企业邮局！</font><br>杂志群发接收对象为订阅了电子杂志的会员（指定邮箱除外）</p>
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="2" height="20" background="images/mmto.gif">
      <p align="center"></p>
    </td>
  </tr>
  <tr>
    <td width="14%" height="26" align="right">收信邮箱：</td>
    <td width="86%" height="26"> <input type="text" name="tomail" style="border:1px black solid" size="36" value="<%=Request.QueryString("mail")%>">&nbsp;                                                                    
      是否群发？<input type="checkbox" name="qunfa" value="ON" <%if Request.QueryString("mail") = "" then Response.Write "checked"%>>                                                                    
    （如果不指定信箱，则默认向订阅杂志的会员信箱群发邮件）</td>                                                                   
  </tr>                                                                   
  <tr>                                                                   
    <td width="14%" height="26" align="right">邮件主题：</td>                                                                   
    <td width="86%" height="26"><input type="text" name="mailsubject" style="border:1px black solid" size="36">                                                                    
    </td>                                                                   
  </tr>                                                                   
  <tr>                                                                   
    <td width="14%" height="26" align="right">邮件格式：</td>                                                                   
    <td width="86%" height="26"><input type="radio" name="radiobutton" value="text" checked  style="border:1px black solid">                                                                   
      文本邮件&nbsp; <input type="radio" name="radiobutton" value="html" style="border:1px black solid">                                                                          
        HTML邮件&nbsp;<font color="#ff0000"> 下面两个内容框要根据选择的邮件格式对应填写，不能同时填写！</font></td>                                                                           
  </tr>                                                                           
  <tr>                                                                           
    <td width="14%" height="26" align="right">邮件内容：<br>（文本格式）</td>                                                                          
    <td width="86%" height="26"><textarea name="mailbody" cols="90" rows="15" style="border:1px black solid"></textarea>  </td>                                                                          
  </tr> 
  <tr>                                                                           
    <td width="14%" height="26" align="right">邮件内容；<br>（HTML格式）</td>                                                                          
    <td width="86%" height="26"><textarea id="editor" name="mailbodyhtm"  style="width:550px;height:280px;visibility:visible"></textarea></td>                                                                          
  </tr>   
  <tr>                                                                          
    <td width="14%" height="26" align="right">邮件等级：</td>                                                                          
    <td width="86%" height="26"><input type="radio" name="dengji" value="1" style="border:1px black solid">                                                                          
      加急&nbsp; <input type="radio" name="dengji" value="3" checked style="border:1px black solid">                                                                          
      普通&nbsp; <input type="radio" name="dengji" value="5" style="border:1px black solid">                                                                          
      低级&nbsp;&nbsp;</td>                                                                          
  </tr>                                                                          
  <tr>                                                                          
    <td width="100%" height="26" colspan="2">                                                                          
      <p align="center"><input type="submit" name="send" value="发送" style="border-style: solid; border-width: 1">                                                                           
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" name="Submit" value="取消" style="border-style: solid; border-width: 1">     <input type=hidden name="action" value="send">                                                                     
    </td>                                                                          
  </tr>                                                                          
 </table></form> 
 <%end Sub%>
</center>                                                                              
</body>                                                                              
</html>
<!--#include virtual="/Include/bottom_superadmin.asp" -->