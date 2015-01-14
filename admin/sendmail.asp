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
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<script language=javascript>
<!--
function CheckAll(form){
for (var i=0;i<form.elements.length;i++){
var e = form.elements[i];
if (e.name != 'chkall') e.checked = form.chkall.checked; 
}
}
-->
</script>
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
<style type="text/css">
/*提示信息*/
.info {position:relative;background:#;color:#666; text-decoration:none;font-size:12px;width:0px;height:0px;text-align:center;border:0px solid #ccc;line-height:18px;}/*设置链接的属性,一定要设置为relative才能使提示层跟着链接走*/
.info:hover {background:#eee;color:#333;}
.info span {display: none }/*设置正常下的span为隐藏状态*/
.info:hover span /*设置hover下的span属性为呈现状态,并设置提示层的位置*/{display:block;position:absolute;top:0px;left:25px;width:200px;
border:1px solid #DC5826; background:#FFF3B4; color:#333;padding:5px;text-align:left;}
</style>
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
	if (document.email.email.value=="" )
	{
	  alert("请填邮箱")
	  document.email.email.focus()
	  return false
	 }
    if((!checkemail(email.email.value))&&(document.email.email.value!=""))
	{
    alert("您输入Email地址不正确，请重新输入!");
    document.email.email.focus();
    return false;
    }
if (document.email.mail_subject.value=="" )
	{
	  alert("请填邮件主题")
	  document.email.mail_subject.focus()
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

<%dim action,email,aJMail,selectedid,i,thisEmail,sssstemp,C_M_Chk
Set aJMail=Server.CreateObject("JMail.Message")   
If aJMail Is Nothing Then  
response.write "<script LANGUAGE='javascript'>alert('温馨小提示：服务器不支持JMail.Message组件，请立即安装此组件，此组件用于邮件发送!');history.go(-1);</script>"
Response.End
End   If
If trim(request("userqf"))="ok" And trim(request("selectedid"))="" Then
response.write "<script LANGUAGE='javascript'>alert('温馨小提示：未选择会员不知邮件要发给谁!');history.go(-1);</script>"
Response.End
End   If
action=request("action")
if action="send" Then
call sendmail()
else
email=request.ServerVariables("query_string")'来自杂志管理
selectedid=trim(request("selectedid"))
If selectedid<>"" Then'来自会员管理
selectedid=split(selectedid,",")
Call OpenConn
for i=0 to ubound(selectedid)
set rs=server.createobject("adodb.recordset")
sql="select email from [DNJY_user] where username='"&trim(cstr(selectedid(i)))&"'"
rs.open sql,conn,1,1
thisEmail=rs(0)
if instr(sssstemp,thisEmail)=0 Then
sssstemp=sssstemp&thisEmail
If i<>ubound(selectedid) Then sssstemp=sssstemp&";"
end If
Next
email=sssstemp
Call CloseRs(rs)
Call CloseDB(conn)
End if%>
<form action="sendmail.asp" method=post name="email" onSubmit="return formcheck()">
<table width="100%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px" bordercolor="#333333" cellspacing="0" cellpadding="2" style="word-break:break-all">
<tr><td colspan=4 bgcolor="#BDBEDE" height=28 align=center><B>向会员发送电子杂志</b></td></tr>
<tr><td width=20% align=center colspan=2><font color="#ff0000"> 注意：由于很多邮局、尤其是免费邮箱群发邮件都受监测，可能因过密或内容相同而被暂时或永久禁止发信功能，如果需经常群发邮件的建议购买不受限制的企业邮局！</font></td></TR>
<tr><td width=20%  align="right">收 件 箱：</td>
<td width="650"><TEXTAREA NAME="email" rows=6 cols="90" style="width=550 ; overflow:auto;" cols="50"><%=email%></TEXTAREA><br><a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>填写收件人邮箱地址多个邮箱以半角逗号或分号隔开千万不要发垃圾邮件哦~后果很严重</span></a> </td></TR>
<tr><td width="20%"  align="right">邮件主题：</td><td width="600"><input type=text name="mail_subject" style="width=90%;" maxlength=50 size="70"><br> <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>填写邮件主题技巧：好的主题吸引人点击阅读差的主题可能会被直接删除</span></a></td></TR>
	<tr>                                                                   
    <td width=20%  align="right">邮件格式：</td>                                                                   
    <td width="86%" height="26"><input type="radio" name="radiobutton" value="text" checked style="border:1px black solid">                                                                   
      文本邮件&nbsp; <input type="radio" name="radiobutton" value="html" style="border:1px black solid">                                                                          
        HTML邮件&nbsp;&nbsp;<font color="#ff0000"> 下面两个内容框要根据选择的邮件格式对应填写，不能同时填写！</font></td>                                                                           
  </tr> 
<tr><td width=20%  align="right">文本格式邮件内容：<br><br><font color=red>（1000字符以内）</font></td><td><TEXTAREA NAME="mailbody" ROWS="15" cols="75"style="width=550px ; overflow:auto;"></TEXTAREA><br> <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>支持文本格式<br>若收件人的邮箱不支持html可选择此方式</span></a> </td></TR>
<tr><td width=20%  align="right">HTML格式邮件内容：<br><br><font color=red>（1000字符以内）</font></td><td>
<textarea id="editor" name="mailbodyhtm"  style="width:550px;height:280px;visibility:visible"></textarea>
&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>支持html但若收件人的邮箱不支持html则看到的是一堆代码而矣</span></a></td></TR>
 
  <tr>                                                                          
    <td width=20%  align="right">邮件等级：</td>                                                                          
    <td width="86%" height="26"><input type="radio" name="dengji" value="1" style="border:1px black solid">                                                                          
      加急&nbsp; <input type="radio" name="dengji" value="3" checked style="border:1px black solid">                                                                          
      普通&nbsp; <input type="radio" name="dengji" value="5" style="border:1px black solid">                                                                          
      低级&nbsp;&nbsp;</td>                                                                          
  </tr>
<tr><td colspan=2 bgcolor='#DADAE9' align=center>
<input type=hidden name=action value="send">
<input type=submit value="立即发送" >
&nbsp; &nbsp; &nbsp; <input type="reset" value="重新填写">
</table>
</form>
<%dim mail_subject,mailbody,mailserver,mailpassword,sitename
end if
sub sendmail()
Dim x,io,sendok,i
If Request("mailbody") <> "" and Request("mailbodyhtm") <> "" Then  
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
mail_subject=trim(request("mail_subject"))
If Request("mailbody") <> "" then
mailbody=Request("mailbody")
Else
mailbody=Request("mailbodyhtm")
End if
email=trim(request("email"))
email=replace(email," ","")
email=replace(email,"；",";")
email=replace(email,"，",";")
email=replace(email,",",";")
if email="" or mail_subject="" or mailbody="" then
response.write "<script language='javascript'>"
response.write "alert('出错了，收件人、邮件主题、邮件内容必须完整填写。');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

if UBound(split(email,";"))>20 then
response.write "<script language='javascript'>"
response.write "alert('出错了，您填写的收件人太多了,限20个！~~');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

mailserver=mailsmtp		
mailname=mailname			
mailpassword=mailpass
sitename=webname
mailbody=mailbody+"◆欢迎光临"+webname+"http://"+web
x=split(email,";")

Dim aJMail   
Set aJMail=Server.CreateObject("JMail.Message")   
If aJMail Is Nothing Then  
response.write "<script LANGUAGE='javascript'>alert('温馨小提示：服务器不支持JMail.Message组件，请立即安装此组件，此组件用于邮件发送!');history.go(-1);</script>"
Response.End
End   If

Dim JMail,ErrStr
for io=0 to Ubound(x)
If ChkMail(x(io)) =true Then'判断邮箱是否合法
'jmail邮件组件发送邮件开始，下面的值要先取得
Set jmail = Server.CreateObject("JMail.Message") 
jmail.silent = true '屏蔽例外错误，返回FALSE跟TRUE两值
jmail.Charset = "gb2312"  '邮件的文字编码为国标
If Request.form("radiobutton") = "html" Then
jmail.ContentType = "text/html"    '邮件的格式为HTML格式
end if
jmail.MailServerUserName = mailname '输入smtp服务器验证登陆名 （邮局中任何一个用户的Email地址） 
jmail.MailServerPassword = mailpassword '输入smtp服务器验证密码 （用户Email帐号对应的密码） 
jmail.From = mailform '发信人信箱，因服务器反垃圾邮件，此信箱必须与发信服务器同域用户的信箱！ 
jmail.FromName = webname '发件人姓名 
jmail.AddRecipient(x(io)) '收件人Email 
jmail.Subject = mail_subject '邮件的标题 
jmail.Body = mailbody '邮件的内容
JMail.Priority =1' Request.form("dengji")      '邮件的紧急程序 
jmail.Send (mailserver) '执行邮件发送smtp服务器地址（企业邮局地址）
ErrStr =jmail.errormessage
jmail.close 
set jmail = nothing '关闭对象
end if'判断邮箱是否合法
next  

if ErrStr="" Then
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
response.write "<script language='javascript'>alert('恭喜，邮件已经发出。');location.href='sendmail_user.asp';</script>"
Response.End
else
Response.Write "<script>alert('发送邮件失败！\n\n失败原因："&ErrStr&"！\n\n点击确定返回！');history.back();</Script>"
Response.End 
end if
end sub
%>
<!--#include virtual="/Include/bottom_superadmin.asp" -->