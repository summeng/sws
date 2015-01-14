
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
<title>不良信息举报</title>
<link rel="stylesheet" href="/<%=strInstallDir%>css/style.css">
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

<script language="javascript">
<!--
//验证邮箱正确性
function checkemail(Bad_infomail){
var str=Bad_infomail;
//在JavaScript中，正则表达式只能使用"/"开头和结束，不能使用双引号
var Expression=/\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*/; 
var objExp=new RegExp(Expression);
if(objExp.test(str)==true){
return true;
}else{
return false;
}
}


/*****  单选按钮检测 此段不用修改 与下面配合*****/
function check_radio(check_radio)
{
    for(i=0;i<check_radio.length;i++){
            if(check_radio[i].checked==true){
                return true;
            }
        }    
    return false;
}
/*****  单选按钮检测end  *****/

//表单判断
function chkinput(){

//下面7行为单选按钮检测，与上面配合
var form1 = document.getElementById("form");
if (!check_radio(form.Bad_infolx))
{
alert("请正确选择举报类型！");
form.Bad_infolx[0].focus();
return false;
}

  if ((!checkemail(form.Bad_infomail.value)) && (document.form.Bad_infomail.value!="" ))
	{
    alert("您输入Email地址不正确，请重新输入!");
    document.form.Bad_infomail.focus();
    return false;
    }
if (document.form.validate_codee.value==0)
	{
	  alert("请填验证码")
	  document.form.validate_codee.focus()
	  return false
	 }
}
// --></script>
</head>
<body>
<TABLE width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
<TR>
	<TD align="center"><b><font color="#FF0000">尊重事实，客观对待</font></b></TD>
</TR>
</TABLE>
<TABLE width="80%" border="0" align="center" cellpadding="0" cellspacing="1">
<TR>
	<TD><b><font color="#FF0000"><b>举报类型(必选)：</b><p></TD></TR>
<form name=form action=Bad_info_chk.asp  method=post onsubmit="return chkinput()">
<TR>
	<TD>
<input type="radio" name="Bad_infolx" value="信息有虚假嫌疑"> 信息有虚假嫌疑<br>
<input type="radio" name="Bad_infolx" value="信息有侵权违法字"> 信息有侵权违法字符<br>
<input type="radio" name="Bad_infolx" value="不按承诺条件交易"> 不按承诺条件交易<br>
<input type="radio" name="Bad_infolx" value="骗取他人钱财"> 骗取他人钱财<br> 
<input type="radio" name="Bad_infolx" value="信息涉及物品假冒伪劣"> 信息涉及物品假冒伪劣<br> 
<input type="radio" name="Bad_infolx" value="信息涉及物品为违禁物品"> 信息涉及物品为违禁物品<br>
<input type="radio" name="Bad_infolx" value="联络信息不对"> 联络信息不对<br>
<input type="radio" name="Bad_infolx" value="其他原因"> 其他原因<p> 

<b>补充说明(选填)：</b><p>
请尽量详细反映实际情况,以便我们及时正确的处理你们的举报<br>

<textarea name="Bad_info"  cols="50" rows="5" class="input2" onkeyUp="textLimitCheck(this, 200);"></textarea><br>
限 200 个字符 已输入 <font color="#CC0000"><span id="messageCount">0</span></font> 个字<p>

<b>您的邮箱(选填)：</b> <input type="text" name="Bad_infomail" value="" size="30" maxLength=50><p>
算式验证：<img src="../tt_getcodee.asp" alt="验证码,看不清楚?请点击刷新验证码" height="10" style="cursor : pointer;" onClick="this.src='../tt_getcodee.asp?t='+(new Date().getTime());" /><input type="text" name="validate_codee" size="4" onkeyup="value=value.replace(/[^\d|^\.]/g,'')"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))">&nbsp;&nbsp; <img src=../images/memo.gif alt="验证码,看不清楚?请点击刷新验证码">
<input type="hidden" name="infoid" value="<%=Request.QueryString("id")%>">
<input type="hidden" name="city_oneid" value="<%=request("city_oneid")%>">
<input type="hidden" name="city_twoid" value="<%=request("city_twoid")%>">
<input type="hidden" name="city_threeid" value="<%=request("city_threeid")%>">
</TD>
</TR>
</TABLE>
	  <TABLE width="80%" border="0" align="center" cellpadding="0" cellspacing="1">
          <tr>
            <td height="35"><div align="center">
              <input name="AddAD" type="submit" class="input1" value="确定举报">
            </div></td>
            <td height="35"><div align="center">
              <input name="AddAD" type="reset" class="input1" value="取消举报">
            </div></td>
          </tr>
        </table>
  </form>
<TABLE width="80%" border="0" align="center" cellpadding="0" cellspacing="1">
             <tr>
                <td height="6" bgcolor="#FBC831">
您每天对同一信息只能举报一次<br>
我们会尽快根据您的举报对信息源进行核实,谢谢您对本站的支持!<br>
本站对举报的信息只作一般处理，不对处理后的信息真实性和合法性负责。</font></td>
                </tr>
</TABLE>

</body>
</html>