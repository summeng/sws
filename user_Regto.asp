<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->
<!--#include file="Include/err.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%if not ChkPost() then 
response.write "禁止站外提交或访问"
response.end 
end If
'if request.cookies("reg")("regkey")="" Or Instr(Request.ServerVariables("HTTP_REFERER"),""&reg&"")=0 then 
'response.redirect ""&reg&"?"&C_ID&"" 
'end If

randomize
session("sms_code")= CInt(8999 * Rnd + 1000)'产生手机验证随机码 

if lmkg2="1" then
call errdick(50)
response.end
end If

Dim tjname
tjname=trim(request("tjname"))%>
<link rel="stylesheet" href="/<%=strInstallDir%>css/css.css" type="text/css">
<style type="text/css">
/*提示信息*/
.info {position:relative;background:#;color:#666; text-decoration:none;font-size:12px;width:0px;height:0px;text-align:center;border:0px solid #ccc;line-height:18px;}/*设置链接的属性,一定要设置为relative才能使提示层跟着链接走*/
.info:hover {background:#eee;color:#333;}
.info span {display: none }/*设置正常下的span为隐藏状态*/
.info:hover span /*设置hover下的span属性为呈现状态,并设置提示层的位置*/{display:block;position:absolute;top:0px;left:25px;width:200px;
border:1px solid #DC5826; background:#FFF3B4; color:#333;padding:5px;text-align:left;}
</style>
<script language="javascript" type="text/javascript">
//用于注册用户名是否重复判断
var xmlHttp = false;
try {
  xmlHttp = new ActiveXObject("Msxml2.XMLHTTP");
} catch (e) {
  try {
    xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
  } catch (e2) {
    xmlHttp = false;
  }
}
if (!xmlHttp && typeof XMLHttpRequest != 'undefined') {
  xmlHttp = new XMLHttpRequest();
}

function callServer() {
//只需填下一行document.getElementById("username").value;中的username为表单中的用户名即可
  var username = document.getElementById("username").value;
  if ((username == null) || (username == "")) return;

var reId=/^[\w\u0391]+$/;//目前为支持大小写字母、数字和下划线字符，如果加\uFFE5为汉字可以使用，如var reId=/^[\w\u0391-\uFFE5]+$/;
var b_id=reId.test(username);
if(!b_id){
var url = "user_regxml.asp?name=" + escape("yes");
xmlHttp.open("GET", url, true);
xmlHttp.onreadystatechange = updatePage
xmlHttp.send(null);
}
	else
  var url = "user_regxml.asp?name=" + escape(username);
  xmlHttp.open("GET", url, true);
  xmlHttp.onreadystatechange = updatePage;
  xmlHttp.send(null);  
}

function updatePage() {
  if (xmlHttp.readyState < 4) {
	test.innerHTML="loading...";
  }
  if (xmlHttp.readyState == 4) {
    var response = xmlHttp.responseText;
	test.innerHTML=response;
  }

}


</script>
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
//验证会员帐号名正确性
function checkeid(username){
var str=username;
//在JavaScript中，正则表达式只能使用"/"开头和结束，不能使用双引号
var Expression=/^[\w\u0391]+$/; //支持大小写字母、数字和下划线字符
var objExp=new RegExp(Expression);
if(objExp.test(str)==true){
return true;
}else{
return false;
}
} 

//验证身份证正确性
function checkeNO(NO){
var str=NO;
//在JavaScript中，正则表达式只能使用"/"开头和结束，不能使用双引号
var Expression=/\d{17}[\d|X]|\d{15}/; 
var objExp=new RegExp(Expression);
if(objExp.test(str)==true){
return true;
}else{
return false;
}
} 


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
    if (!checkeid(thisForm.username.value))
	{
    alert("登录帐号仅支持大小写字母、数字和下划线字符，请重新输入!");
    document.thisForm.username.focus();
    return false;
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
    if (document.thisForm.mobile.value=="" )
	{	  
      alert("手机号码不能为空，请重新输入！");
	  document.thisForm.mobile.focus()
	  return false
	 }
	var sMobile = document.thisForm.mobile.value
	if((document.thisForm.mobile.value!="") && (!(/^((\(\d{3}\))|(\d{3}\-))?13[0-9]\d{8}|14[0-9]\d{8}|15[0-9]\d{8}|17[0-9]\d{8}|18[0-9]\d{8}/.test(sMobile)))) //(新添加了15,18,17两个号段)
	{
		alert("不是完整的11位手机号或者正确的手机号前七位");
		document.thisForm.mobile.focus();
		return false;
	}
 <%if sms_kg=True then%>
	
    if (document.thisForm.captcha.value=="" )
	{	  
      alert("手机短信校验码不能为空，请重新输入！");
	  document.thisForm.captcha.focus()
	  return false
	 }
    if(document.thisForm.captcha.value != <%=session("sms_code")%>) 
	{	  
      alert("手机短信校验码与系统生成的不符，若时间太长失效请联系客服删除记录后重新获取手机短信校验码！");
	  document.thisForm.captcha.focus()
	  return false
	 }
 <%end if%>
<%If mailqr=1 then%>
   if(document.thisForm.email.value.length<1)
	{
	    alert("邮箱地址不能为空!");
		document.thisForm.email.focus()
	    return false;
	}
 <%end if%>
    if ((document.thisForm.email.value!="")&&(!checkemail(thisForm.email.value)))
	{
    alert("您输入Email地址不正确，请重新输入!");
    document.thisForm.email.focus();
    return false;
    }
   if(document.thisForm.email.value.length>30 )
	{
	    alert("信箱不能超过30个字符!");
	    document.thisForm.email.focus();
	    return false;
	}

 <%if codekg1=1 then%>
    if (document.thisForm.wenti.value=="" )
	{	  
      alert("验证答案不能为空，请重新输入！");
	  document.thisForm.wenti.focus()
	  return false
	 }
	
    if(document.thisForm.wenti.value != document.thisForm.daan.value) 
	{	  
      alert("验证答案不对，请重新输入！");
	  document.thisForm.wenti.focus()
	  return false
	 }
 <%end if%>
 <%if codekg2=1 then%>
    if (document.thisForm.yzm.value=="" )
	{	  
      alert("数字验证码不能为空，请重新输入！");
	  document.thisForm.yzm.focus()
	  return false
	 }
<%end if%>
 <%if codekg3=1 then%>
    if (document.thisForm.code.value=="" )
	{	  
      alert("文字验证码不能为空，请重新输入！");
	  document.thisForm.code.focus()
	  return false
	 }
<%end if%>
<%if codekg4=1 then%>
    if (document.thisForm.ttdv.value=="" )
	{	  
      alert("验证星期不能为空，请重新输入！");
	  document.thisForm.ttdv.focus()
	  return false
	 }
<%end if%>
<%if codekg5=1 then%>
    if (document.thisForm.validate_codee.value=="" )
	{	  
      alert("算式验证码不能为空，请重新输入！");
	  document.thisForm.validate_codee.focus()
	  return false
	 }
<%end if%>
	if((document.thisForm.answer1.value=="")&&(document.thisForm.problem.value!=""))
        {
            alert("已填取回密码问题就要填答案!");
			document.thisForm.answer1.focus()
            return false;
        }
    if((document.thisForm.problem.value.length>=1) && (document.thisForm.problem.value == document.thisForm.answer1.value)) {
	document.thisForm.answer1.focus();
	document.thisForm.answer1.value = '';
	document.thisForm.answer2.value = '';
    alert("密码问题与答案不要相同，请重新输入！");
	return false;
  }

	   if(document.thisForm.answer1.value!=document.thisForm.answer2.value)
        {
            alert("两次输入答案不一致!");
			document.thisForm.answer2.focus()
            return false;
        }
    if ((!checkeNO(thisForm.idcard.value)) && (document.thisForm.idcard.value!=""))
   {
    alert("您输入身份证号码不正确!");
    document.thisForm.idcard.focus();
    return false;
    }
//var sMobile = document.thisForm.dianhua.value
//if((document.thisForm.dianhua.value!="") && (!(/^([0\+]\d{2,3}-)?(0[0-9]{2,3}\-)?([2-9][0-9]{6,7})+(\-[0-9]{1,4})?$/.test(sMobile)))) //(国家代号-区号-电话-分机格式(可选)：+086-010-85858585-2538)
//	{
//		alert("不规范的电话号码");
//		document.thisForm.dianhua.focus();
//		return false;
//	}
if(document.thisForm.dianhua.value.length>30 )
	{
	    alert("固定电话不能超过30个字符!");
		document.thisForm.dianhua.focus();
	    return false;
	}

//var sMobile = document.thisForm.fax.value
//if((document.thisForm.fax.value!="") && (!(/^([0\+]\d{2,3}-)?(0[0-9]{2,3}\-)?([2-9][0-9]{6,7})+(\-[0-9]{1,4})?$/.test(sMobile)))) //(国家代号-区号-电话-分机格式(可选)：+086-010-85858585-2538)
//	{
//		alert("不规范的传真电话号码");
//		document.thisForm.fax.focus();
//		return false;
//	}
if(document.thisForm.fax.value.length>30 )
	{
	    alert("传真不能超过30个字符!");
		document.thisForm.fax.focus();
	    return false;
	}

 if (document.thisForm.http.value=="http://" )     
  {   
  alert("网址前面不要带http://")
  document.thisForm.http.focus();
  return false;  
  }   
  
}
//-->
</SCRIPT>
<!--文字验证码调用开始-->
<SCRIPT LANGUAGE=javascript>
/*显示认证码 o start1*/
function get_Code() {
        var Dv_CodeFile = "Dv_GetCode.asp";
        if(document.getElementById("imgid"))
                document.getElementById("imgid").innerHTML = '<img src="'+Dv_CodeFile+'?t='+Math.random()+'" alt="点击刷新验证码" style="cursor:pointer;border:0;vertical-align:middle;height:30px;" onclick="this.src=\''+Dv_CodeFile+'?t=\'+Math.random()" />'
}
/*o end*/
</script>
<script language="JavaScript" type="text/javascript">
var dvajax_request_type = "GET";
</script>
<script language="JavaScript" src="dv_ajax.js" type="text/javascript"></script>
<!---文字验证码调用结束-->
<!---手机短信校验码调用开始-->
<script language="javascript">
	function get_mobile_code(){
//判断是否输入手机号
        var myReg = /^1[3,4,5,7,8]\d{9}$/g;
        var mobile = document.thisForm.mobile.value;
        if (!myReg.exec(mobile)) {
            alert("请输入11位正确的手机号码！" + mobile);
            document.thisForm.mobile.focus();
            return false;
        }
//发送校验码
        $.post('sms.asp', {mobile:jQuery.trim($('#mobile').val())}, function(msg) {
            alert(jQuery.trim(unescape(msg)));
        });
	};
</script>
<script type="text/javascript" src="Include/sms_jQuery-1.6.js"></script>
<script type="text/javascript" src="Include/sms_button.js"></script>
<!---手机短信验证调用结束-->
<script language="javascript" type="text/javascript">
<!--
	function checkPwd()
	{
	}
	first("selectp","selectc","thisForm",0,0);
	function showadv()
	{
		if (document.thisForm.advshow.checked == true) 
		{
			document.getElementById("adv").style.display = "inline";
			document.getElementById("advance").innerHTML="关闭高级设置选项";
		}
		else
		{
			document.getElementById("adv").style.display = "none";
			document.getElementById("advance").innerHTML="显示高级设置选项";
		}
	}
//-->
</script> 
</HEAD>
<center>
<TABLE width=980 border=0 cellpadding=3 cellspacing=1 align=center  bgcolor=#FFFFFF    class="grayline"  background="/img/hd_normal_bg3.gif">
<tr><td height=30>目前位置：<a href=index.asp?<%=C_ID%>>首页</a> > 会员注册第二步</td></tr>
</TABLE>
<TABLE width=980 border=0 cellpadding=3 cellspacing=1 align=center  bgcolor=#FFFFFF height=550    class="grayline" >
<tr><td height="1" background="images/bgline.gif"></td></tr>
<tr><td>
<body topmargin="0" leftmargin="0">
<div>
  <center>
  <table width="980" height="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" bgcolor="#B6E688" style="border-collapse: collapse">
    <tr>
      <td height="156"><div >
        
        <br><table width="800" align=center border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">    
                <tr>                 
                  <td width="800" bgcolor="#FFFFFF">
                    <table width="800" height="126" border="1" cellpadding="3" cellspacing="3" bordercolor="#CCCCCC">
					
                   <form name="thisForm" method="post" action="user_reg_save.asp?<%=C_ID%>" language="javascript" onsubmit="return CheckForm()">   
                      <tr bgcolor="#FFFFFF"> 
                        <td width="654" height="21" colspan=2>
                        <p align="CENTER"><strong>为确保信息真实有效，请耐心填写注册信息 (带<font color="#ff0000">*</font>必填)</strong></td>
                      </tr>
		             <tr bgcolor="#6CCEE5"> 
                        <td height="30" align="center" colspan="2">
                        <font color="#ff0000">必 填 项</font></td>
                    </tr>
                   <tr bgcolor="#FFFFFF"> 
                      <td height="30">
                      <p align="right">登陆帐号：</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" id="username" maxlength="16" name="username" size="20" onpropertychange="kn()" autocomplete="off" onChange="callServer();"><font color="#ff0000"> *</font> &nbsp;<span id="test"><img src='images/reg_warning.gif'>请输入帐号检测有效性</span><br>&nbsp;&nbsp;<span id=num>你还可以输入16个字符</span>&nbsp;&nbsp;(仅允许英文和字数)</td>
                    </tr>
 
                      </tr>
                   <tr bgcolor="#FFFFFF"> 
                      <td height="30">
                      <p align="right">登录密码：</td>
                      <td height="30">
                      &nbsp;
                        <input class="inputa" type="password" name="password" size="20" onpropertychange="kn2()"><font color="#ff0000"> *</font> <span id=num2>5-12位,你还可以输入12个字符</span></td>
                      </tr>
                      </tr>
                   <tr bgcolor="#FFFFFF"> 
                      <td height="30">
                      <p align="right">确认密码：</td>
                      <td height="30">
                      &nbsp;
                        <input class="inputa" type="password" name="password1" size="20" ><font color="#ff0000"> *</font> &nbsp;&nbsp;&nbsp;</td>
                      </tr>
                      </tr>

 
                    <tr bgcolor="#FFFFFF"> 
                      <td height="30">
                      <p align="right">性&nbsp;&nbsp;&nbsp;&nbsp;别：</td>
                      <td height="30">
                      &nbsp;
                      <select id="ctlSex" name="Sex" class="inputa">
                        <option value="1" selected>男</option>
                        <option value="0">女</option>
                      </select>  <font color="#ff0000"> *</font></td>
                    </tr>

                     <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      手机号码：</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="30" id="mobile" name="mobile" size="24" onKeyUp="if(isNaN(value)){alert('移动电话只允许输入数字');value='';}"> <font color="#ff0000"> *</font> </td>
                    </tr>
					<%if sms_kg=True then%>
                     <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      校 验 码：</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" size="8" name="captcha" class="inputBg" /> <font color="#ff0000"> *</font> 查看手机短信校验码并填在此 <input id="zphone" type="button" value=" 获取手机校验码 " onclick="get_mobile_code();countDown(this,30);" ></td>
                    </tr>  
					<%End if%>
          			<tr bgcolor="#FFFFFF"> 
                        <td height="30" align="right">
                        电子信箱：</td>
                        <td height="13">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="255" name="email" size="24" > <%If mailqr=1 then%><font color="#ff0000"> *</font>  注册本站会员需邮件验证，<%End if%>请填正确信箱！</td>
                    </tr>
                   <tr bgcolor="#FFFFFF"> 
                      <td height="30">
                      <p align="right">推 荐 人：</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="16" name="recommend" value="<%=tjname%>" size="24"> 填写推荐人会员ID号可给其增加积分 <%=jf_5%> 分，没有请留空。</td>
                    </tr>

<%Call OpenConn
if codekg1=1 then
	  '问答式验证
dim rscheck
Randomize 
Set rscheck= Server.CreateObject("ADODB.Recordset")
If SystemDatabaseType = 1 Then
sql="select  * from DNJY_wenda where xs=1 order BY newid()"
Else
sql="select  * from DNJY_wenda where xs=1 order BY rnd(-(ID+"& rnd() &"))"
End if
rscheck.open sql,conn,1,1
%>
                     <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      问答验证：</td>
                      <td height="30">
                      &nbsp;问题：<%=rscheck("wenti")%><br>&nbsp;答案：<input type="text" name="wenti" size="12" class="inputa"> <font color="#ff0000">*</font></td>
                    </tr><input type="hidden" name="daan" value="<%=rscheck("daan")%>">
	<%rscheck.close
	set rscheck=Nothing
	End If
	if codekg2=1  then%>
                     <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      数字验证：</td>
                      <td height="30">
                      &nbsp;<input type="text" name="yzm" size="4" class="inputa"/><img src="tt_getcode.asp" height="18" alt="验证码,看不清楚?请点击刷新验证码" style="cursor : pointer;" onclick="this.src='tt_getcode.asp'"/>&nbsp;<font color="#ff0000">*</font>&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>验证码,看不清楚?请点击刷新验证码</span></a></td>
                    </tr>
	  <%End if
	if codekg3=1 then%>
                     <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      文字验证：</td>
                      <td height="30">
                      &nbsp;<input type="text" class="inputa" name="code" id="code" size="4" maxlength="4" tabindex="6" onfocus="get_Code();this.onfocus=null;" onkeyup="dv_ajaxcheck('checke_dvcode','code');"  autocomplete="off"/>
    <font color="#ff0000">*</font> <span id="imgid" style="color:red">点击获取验证码</span><span id="isok_code"></span></td>
                    </tr>
<%End if%>
<%if codekg4=1 then%>
                     <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      星期验证：</td>
                      <td height="30">
                      &nbsp今天是 <select name="ttdv" class="inputa">
<option value="" selected="selected">请选择</option>
<option value="1">星期一</option>
<option value="2">星期二</option>
<option value="3">星期三</option>
<option value="4">星期四</option>
<option value="5">星期五</option>
<option value="6">星期六</option>
<option value="0">星期日</option>
</select>&nbsp;&nbsp; <font color="#ff0000">*</font><a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>请选择正确的星期几</span></a></td>
      </tr>
<%End if%>
<%if codekg5=1 then%>
                     <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      算式验证：</td>
                      <td height="30">
                      &nbsp;<img src="tt_getcodee.asp" alt="验证码,看不清楚?请点击刷新验证码" height="10" style="cursor : pointer;" onClick="this.src='tt_getcodee.asp?t='+(new Date().getTime());" /><input name="validate_codee" class="inputa" type="text" tabindex="3" value="" size="4" maxlength="4" onkeyup="value=value.replace(/[^\d|^\.]/g,'')"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))"/> &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>验证码,看不清楚?请点击刷新验证码</span></a></td>
                    </tr>
<%End if%>

 			<tr bgcolor="#6CCEE5"> 
                        <td height="30" align="center" colspan="2">
                        <font color="#ff0000"><INPUT id=advcheck name=advshow type=checkbox value=1 onclick=showadv()>
         <span id=advance>显示高级设置选项</span>（建议完善以下资料方便以后发布信息）</font></td>
                    </tr>
</table>
<table style="DISPLAY: none" id=adv width="800" height="126" border="1" cellpadding="3" cellspacing="3" bordercolor="#CCCCCC">
                  <tr bgcolor="#FFFFFF"> 
                      <td height="30" width="88">
                      <p align="right">密码问题：</td>
                      <td height="30">
                      &nbsp;
                        <input class="inputa" type="problem" maxlength="20" name="problem" size="20" > 填写后方便忘记密码找回</td>
                      </tr>
                    
                   <tr bgcolor="#FFFFFF"> 
                      <td height="30">
                      <p align="right">密码答案：</td>
                      <td height="30">
                      &nbsp;
                        <input class="inputa" type="answer1" maxlength="20" name="answer1" size="20" >　找回密码时要回答</td>
                      </tr>                     
                      
                   <tr bgcolor="#FFFFFF"> 
                      <td height="30">
                      <p align="right">确认答案：</td>
                      <td height="30">
                      &nbsp;
                        <input class="inputa" type="answer2" maxlength="20" name="answer2" size="20" > &nbsp;&nbsp;&nbsp;</td>
                      </tr>
                   <tr bgcolor="#FFFFFF"> 
                      <td height="30">
                      <p align="right">真实姓名：</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="12" name="name" size="20"> </td>
                    </tr>

                    <tr bgcolor="#FFFFFF"> 
                      <td height="30">
                      <p align="right">身 份 证：</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="20" name="idcard" size="24" > （不对外公开）</td>
                    </tr>         

       
			<tr bgcolor="#FFFFFF">
			 <td height="30" align="right">QQ&nbsp;号码：</td>
			<td height="14">　<input class="inputa" type="text" maxlength="255" name="qq" size="24" onKeyUp="if(isNaN(value)){alert('QQ号码只允许输入数字');value='';}"></td>
			        </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      微信号码：</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="50" name="weixin" size="24" ></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      公司名称：</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="80" name="mpname" size="37" > 也作企业名片名称</td>
                    </tr>					  
                    <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      通信地址：</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="100" name="dizhi" size="60" ></td>
                    </tr>
                     <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      邮政编码：</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="6" name="youbian" size="12"  onKeyUp="if(isNaN(value)){alert('邮政编码只允许输入数字');value='';}">    </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      固定电话：</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="30" name="dianhua" size="37" > 国家代号-区号-电话-分机格式(可选部分)：+086-010-85858585-2538</td>
                    </tr>
                     <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      传&nbsp;&nbsp;&nbsp;&nbsp;真：</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="30" name="fax" size="37" > 国家代号-区号-电话-分机格式(可选部分)：+086-010-85858585-2538</td>
                    </tr>
                     <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      网&nbsp;&nbsp;&nbsp;&nbsp;址：</td>
                      <td height="30">
                      &nbsp;  
                      <input class="inputa" type="text" value=""  maxlength="100" name="http" size="60"> 范例www.ip126.com，前面不能带http:// </td>
                    </tr>
			<tr bgcolor="#FFFFFF"> 
                        <td height="30" align="right">
                        所属地区：</td>
                        <td height="13">
                      &nbsp;
                      <%
set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")
%>
          <SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
	dim count5:count5 = 0
        do while not rs.eof 
        %>
subcat[<%=count5%>] = new Array("<%=rs("city")%>","<%=rs("id")%>","<%=rs("twoid")%>");
        <%
        count5 = count5 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount=<%=count5%>;
          </SCRIPT>
          <%
set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid>0")
%>
          <SCRIPT language = "JavaScript">
var onecount4;
onecount4=0;
subcat4 = new Array();
        <%
		dim count4:count4 = 0
        do while not rs.eof 
        %>
subcat4[<%=count4%>] = new Array("<%=rs("city")%>","<%=rs("id")%>","<%=rs("twoid")%>","<%=rs("threeid")%>");
        <%
        count4 = count4 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount4=<%=count4%>;

function changelocation(locationid)
    {
    document.thisForm.city_two.length = 0; 
	document.thisForm.city_two.options[0] = new Option('选择城市','');
	document.thisForm.city_three.length = 0; 
	document.thisForm.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.thisForm.city_two.options[document.thisForm.city_two.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.thisForm.city_three.length = 0; 
    document.thisForm.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.thisForm.city_three.options[document.thisForm.city_three.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
          <SELECT name="city_one" size="1" id="select" class="inputa" onChange="changelocation(document.thisForm.city_one.options[document.thisForm.city_one.selectedIndex].value)">
            <OPTION value="" selected>选择地区</OPTION>
            <%set rs=conn.execute("select * from DNJY_city where id>0 and twoid=0 order by indexid asc")
if rs.eof or rs.bof then
response.write "<option value=''>没有分类</option>"
else
do until rs.eof
response.write "<option value='"&rs("id")&"'>"&rs("city")&"</option>"
    rs.movenext
    loop
end if
Call CloseRs(rs)
%>
          </SELECT>
          <SELECT name="city_two" id="select6" class="inputa" onChange="changelocation4(document.thisForm.city_two.options[document.thisForm.city_two.selectedIndex].value,document.thisForm.city_one.options[document.thisForm.city_one.selectedIndex].value)">
            <OPTION value="" selected>选择城市</OPTION>
          </SELECT>
          <SELECT name="city_three" id="select7" class="inputa">
            <OPTION value="" selected>选择城市</OPTION>
          </SELECT> （默认所在地区，填写方便以后发布信息自动填写）</td>
                    </tr>
			<tr bgcolor="#FFFFFF"> 
                        <td height="30" align="right">
                        信息分类：</td>
                        <td height="13">
                      &nbsp;
                      <%
set rs=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid=0")
%>
          <SCRIPT language = "JavaScript">
var onecount6;
onecount6=0;
subcat7 = new Array();
        <%
	dim count6:count6 = 0
        do while not rs.eof 
        %>
subcat7[<%=count6%>] = new Array("<%=rs("name")%>","<%=rs("id")%>","<%=rs("twoid")%>");
        <%
        count6 = count6 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount6=<%=count6%>;
          </SCRIPT>
          <%
set rs=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid>0")
%>
          <SCRIPT language = "JavaScript">
var onecount7;
onecount7=0;
subcat5 = new Array();
        <%
		dim count7:count7 = 0
        do while not rs.eof 
        %>
subcat5[<%=count7%>] = new Array("<%=rs("name")%>","<%=rs("id")%>","<%=rs("twoid")%>","<%=rs("threeid")%>");
        <%
        count7 = count7 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount7=<%=count7%>;

function changelocation6(locationid)
    {
    document.thisForm.type_two.length = 0; 
	document.thisForm.type_two.options[0] = new Option('选择分类','');
	document.thisForm.type_three.length = 0; 
	document.thisForm.type_three.options[0] = new Option('选择分类','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount7; i++)
        {
            if (subcat7[i][1] == locationid)
            { 
                document.thisForm.type_two.options[document.thisForm.type_two.length] = new Option(subcat7[i][0], subcat7[i][2]);
            }        
        }
        
    }    
	
	function changelocation7(locationid,locationid1)
    {
    document.thisForm.type_three.length = 0; 
    document.thisForm.type_three.options[0] = new Option('选择分类','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount7; i++)
        {
            if (subcat5[i][2] == locationid)
            { 
			if (subcat5[i][1] == locationid1)
			{
                document.thisForm.type_three.options[document.thisForm.type_three.length] = new Option(subcat5[i][0], subcat5[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
          <SELECT name="type_one" size="1" id="select" class="inputa" onChange="changelocation6(document.thisForm.type_one.options[document.thisForm.type_one.selectedIndex].value)">
            <OPTION value="" selected>选择分类</OPTION>
            <%set rs=conn.execute("select * from DNJY_type where id>0 and twoid=0 order by indexid asc")
if rs.eof or rs.bof then
response.write "<option value=''>没有分类</option>"
else
do until rs.eof
response.write "<option value='"&rs("id")&"'>"&rs("name")&"</option>"
    rs.movenext
    loop
end if
Call CloseRs(rs)
%>
          </SELECT>
          <SELECT name="type_two" id="select6" class="inputa" onChange="changelocation7(document.thisForm.type_two.options[document.thisForm.type_two.selectedIndex].value,document.thisForm.type_one.options[document.thisForm.type_one.selectedIndex].value)">
            <OPTION value="" selected>选择分类</OPTION>
          </SELECT>
          <SELECT name="type_three" id="select7" class="inputa">
            <OPTION value="" selected>选择分类</OPTION>
          </SELECT> （默认信息分类，填写方便以后发布信息自动填写）</td>
                    </tr>
			      <tr bgcolor="#FFFFFF"> 
                        <td height="30" align="right">
                        订阅资讯：</td>
                        <td height="13"> <input type="radio" name="maillist" value="1" >                 
                  我想订阅&nbsp;&nbsp;&nbsp;&nbsp;                          
                  <input type="radio" name="maillist" value="0" checked>                 
                  不想订阅
                      &nbsp;（订阅后可信箱接收本站发送的资讯）
                      </td>
                    </tr>					
                    </table>
                 
	         <table width="800" align=center border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">		           
                    <tr bgcolor="#FFFFFF"> 
                      <td height="25" colspan="3"><p align="CENTER"> 
					  <input type="submit" value=" 确 定 " name="yes" onclick="javascript:return CheckForm();" language="javascript" id="yes" >
                      <input name="reset" type="reset" value=" 重 写 ">			 
                </form>
                      </td>
                    </tr>
               
              </table>
</td></tr>
</table> 
</center>
</BODY>
</HTML>

<%
If mailqr=1 then
Dim rsdq
set rsdq=server.createobject("adodb.recordset")
sql="delete from [DNJY_usertemp] where DateDiff("&DN_DatePart_D&",zcdata,"&DN_Now&")>"&regyxq&""'删除3天未邮件认证的临时注册信息
rsdq.open sql,conn,1,3
End If
Response.Cookies("reg")("regkey")=""

If sms_del>0 then
set rsdq=server.createobject("adodb.recordset")
sql="delete from [DNJY_sms] where DateDiff("&DN_DatePart_D&",sms_addtime,"&DN_Now&")>"&sms_del&""'删除指定天数后手机短信验证记录
rsdq.open sql,conn,1,3
End If

Call CloseDB(conn)%>