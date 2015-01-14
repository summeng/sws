<!--#include file="../conn.asp"-->
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
<%call checkmanage("04")%>
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
.style1 {	color: #FFFFFF;
	font-weight: bold;
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
//显示输入字符数

function kn()
  {
   var ln=thisForm.username.value.Length();
     num.innerText='您还可以输入:'+(16-ln)+'个字符';
      //if (ln>=16) form.username.readOnly=true;  //这行是如果满足条件表单无法输入和修改
}
function kn2()
 {
   var ln=thisForm.password.value.Length();
     num2.innerText='您还可以输入:'+(12-ln)+'个字符';
      
}
// --></script>
<script language="javascript" type="text/javascript">
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
function CheckForm()
{

        if(document.thisForm.username.value.length<1)
	{
	    alert("登录帐号不能为空!");
		document.thisForm.username.focus()
	    return false;
	}
     if ((document.thisForm.username.value.Length()>16) || (document.thisForm.username.value.Length()<1)) //字节数限制，与上面配合
     {
	 alert("登录帐号长度要求1－16字节，请重新输入！");
	  document.thisForm.username.focus()
	  return false
  }

	  if ((document.thisForm.password.value.Length()>12) || (document.thisForm.password.value.Length()<5)) //字节数限制，与上面配合
     {
	 alert("密码长度要求5－12字节，请重新输入！");
	  document.thisForm.password.focus()
	  return false
  }
	    if(document.thisForm.password1.value.length<1)
	{
	    alert("确认密码不能为空!");
		document.thisForm.password1.focus()
	    return false;
	}
	   if(document.thisForm.password1.value!=document.thisForm.password.value)
        {
            alert("两次输入密码不一致!");
			document.thisForm.password1.focus()
            return false;
        }
	if((!checkemail(thisForm.email.value))&&(document.thisForm.email.value!=""))
    {
    alert("您输入Email地址不正确，请重新输入!");
    document.thisForm.email.focus();
    return false;
    }
	}
//-->
</SCRIPT>

<div align="center">
  <center>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%" height="109">
<form  name="thisForm" method="POST" action="add_user_save.asp"  onsubmit="return CheckForm()">
<TR>
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">直 接 添 加 用 户</FONT></TD>
 </TR>
  <tr bgcolor="#FFFFFF">
    <td height="44" style="border-top-style: solid; border-top-width: 1"><table width="100%"  border="0" cellspacing="6" cellpadding="6">
      <tr>
        <td width="16%">用 户 名：</td>
        <td width="84%"><input type="text" name="username" maxlength="16" size="30"> <font color="#ff0000"> *</font></td>
      </tr>
      <tr>
        <td width="16%">密&nbsp;&nbsp;&nbsp;&nbsp;码：</td>
        <td width="84%"><input type="text" name="password" size="30"> <font color="#ff0000"> *</font></td>
      </tr>
      <tr>
        <td width="16%">确认密码：</td>
        <td width="84%"><input type="text" name="password1" size="30"> <font color="#ff0000"> *</font></td>
      </tr>
      <tr>
        <td>邮件地址：</td>
        <td><input type="text" name="email" size="30"> 建议填写邮箱</td>
      </tr>
            <tr>
        <td>订阅资讯：</td>
        <td><input type="radio" name="maillist" value="1" >                 
                  订阅&nbsp;&nbsp;&nbsp;&nbsp;                          
                  <input type="radio" name="maillist" value="0" checked>                 
                  不订阅</td>
      </tr>
    </table></td>
    </tr>
  <tr bgcolor="#eeeeee">
    <td height="35"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><div align="center">
            <input type="submit" value="确认添加" name="B1">
        </div></td>
        <td><div align="center">
            <input type="reset" value="取消添加" name="B1">
        </div></td>
      </tr>
    </table></td>
    </tr>
  </form>
</table>
  </center>
</div>
