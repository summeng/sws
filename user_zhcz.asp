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
<style type="text/css">
<!--
.style5 {color: #FF0000; font-weight: bold; }
.style6 {color: #FF0000}
-->
</style>
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
<script language="JavaScript">
function CheckForm()
{

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


	
	if (document.alipayment.aliname.value.length == 0) {
		alert("请输入购买人的姓名.");
		document.alipayment.aliname.focus();
		return false;
	}
	
		if (document.alipayment.alimoney.value.length == 0) {
		alert("请输入付款金额.");
		document.alipayment.alimoney.focus();
		return false;
	}
   if (isNaN(document.alipayment.alimoney.value))
	{
	  alert("付款金额应为数字！")
	  document.alipayment.alimoney.focus()
	  return false
	 }

	if(document.alipayment.dianhua.value=="" && document.alipayment.CompPhone.value=="") 
	{
	    alert("固定电话和移动电话不能同时为空!");
		document.alipayment.dianhua.focus()
	    return false;
	}
//var sMobile = document.alipayment.dianhua.value
//if((document.alipayment.dianhua.value!="") && (!(/^([0\+]\d{2,3}-)?(0[0-9]{2,3}\-)?([2-9][0-9]{6,7})+(\-[0-9]{1,4})?$/.test(sMobile)))) //(国家代号-区号-电话-分机格式(可选)：+086-010-85858585-2538)
//	{
//		alert("不规范的电话号码");
//		document.alipayment.dianhua.focus();
//		return false;
//	}	
//	var sMobile = document.alipayment.CompPhone.value
//	if((document.alipayment.CompPhone.value!="") && (!(/^((\(\d{3}\))|(\d{3}\-))?13[0-9]\d{8}|15[0-9]\d{8}|18[0-9]\d{8}/.test(sMobile)))) //(检验13，15，18号段)
//	{
//		alert("不是完整的11位手机号或者正确的手机号前七位");
//		document.alipayment.CompPhone.focus();
//		return false;
//	}
    if((!checkemail(alipayment.email.value))&&(document.alipayment.email.value!=""))
	{
    alert("您输入Email地址不正确，请重新输入!");
    document.alipayment.email.focus();
    return false;
    }
   if(document.alipayment.email.value.length>30 )
	{
	    alert("信箱不能超过30个字符!");
	    document.alipayment.email.focus();
	    return false;
	}


	return true;
}

</script>
</head>

<body topmargin="0" leftmargin="0">
<%
dim email,name,oicq,dianhua,CompPhone,dizhi,youbian
%>

<div align="center">
  <center>
  <table width="980" height="423" border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" bgcolor="#FFFFFF" style="border-collapse: collapse">
    <tr>
      <td width="214" height="351" valign="top"><div align="center">
            <!--#include file=userleft.asp-->　
        </div></td>
    <td width="5">&nbsp;</td>    
<%username=request.cookies("dick")("username")
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_user] where username='"&username&"'" 
rs.open sql,conn,1,1
if rs.eof then
response.write "<br>"
response.write "<li>参数错误！"
response.end
end if
email=rs("email")
name=rs("name")
dianhua=rs("dianhua")
CompPhone=rs("CompPhone")
oicq=rs("qq")
dizhi=rs("dizhi")
youbian=rs("youbian")
Call CloseRs(rs)%>
      <td width="760" align="right" valign="top" bgcolor="#FFFFFF"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="100%" class="ty1">
        <tr bgcolor="#FAFAFA">
          <td height="25"><span class="style5">　帐户充值：</span></td>
        </tr>

        <form name="alipayment" action="user_zhcz_chk.asp?<%=C_ID%>" method="POST"  onSubmit="return CheckForm();">
          <tr>
            <td align="center"><table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
                <tr>
                  <td><TABLE width="100%" height=100 border="0" align=center cellpadding="6" cellspacing="1" style="FONT-SIZE: 9pt">
                      <TBODY>
                   <TR class=tdbg>
                          <TD width="16%" bgcolor="#F8F8F8">您的帐号：</TD>
                          <TD width="86%" bgcolor="#FFFFFF"><b><font color="#0000FF"><%=username%></font></b> </TD>
                        </TR>
                        <TR class=tdbg>
                          <TD width="16%" bgcolor="#F8F8F8">姓&nbsp;&nbsp;&nbsp;&nbsp;名：</TD>
                          <TD width="86%" bgcolor="#FFFFFF"><input type="text" name="aliname" class="form" size="20" value="<%=name%>">
                            （必填）</TD>
                        </TR>
                        <TR class=tdbg>
                          <TD width="16%" bgcolor="#F8F8F8">充值金额：</TD>
                          <TD width="86%" bgcolor="#FFFFFF"><input onKeyPress="return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,/^\d*\.?\d{0,2}$/,&#9;event.dataTransfer.getData('Text'))" maxLength=10 size=20 name=alimoney class="form">
                            元人民币（必填，格式：<span class="style6">50.00</span>）</TD>
                        </TR>

                        <TR class=tdbg>
                          <TD width="16%" bgcolor="#F8F8F8">固定电话：</TD>
                          <TD bgcolor="#FFFFFF"><input type="text" name="dianhua" class="form" size="20" value="<%=dianhua%>">
                            <br> (国家代号-区号-电话-分机格式(可选)：+086-010-85858585-2538)</TD>
                        </TR>
                        <TR class=tdbg>
                          <TD width="16%" bgcolor="#F8F8F8">移动电话：</TD>
                          <TD bgcolor="#FFFFFF"><input type="text" name="CompPhone" class="form" size="20" value="<%=CompPhone%>" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')">
                            （<font color="#CC5200">固定电话和移动电话必填其一</font>）</TD>
                        </TR>                        
                        <TR class=tdbg>
                          <TD width="16%" bgcolor="#F8F8F8">电子信箱：</TD>
                          <TD bgcolor="#FFFFFF"><input type="text" name="email" size="20" value="<%=email%>">
                            </TD>
                        </TR>
                        <TR class=tdbg>
                          <TD width="16%" bgcolor="#F8F8F8">联络QQ号：</TD>
                          <TD bgcolor="#FFFFFF"><input type="text" name="oicq" size="20" value="<%=oicq%>" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')">
                            </TD>
                        </TR>
                        <TR class=tdbg>
                          <TD width="16%" bgcolor="#F8F8F8">联络地址：</TD>
                          <TD bgcolor="#FFFFFF"><input type="text" name="dizhi" size="20" value="<%=dizhi%>">
                            </TD>
                        </TR>
                        <TR class=tdbg>
                          <TD width="16%" bgcolor="#F8F8F8">邮政编码：</TD>
                          <TD bgcolor="#FFFFFF"><input type="text" name="youbian" size="20" value="<%=youbian%>" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')">
                            </TD>
                        </TR>                        
                        <TR class=tdbg>
                          <TD width="16%" bgcolor="#F8F8F8">我要留言：</TD>
                          <TD bgcolor="#FFFFFF">
      <textarea name="alibody" cols="50" rows="3" onkeyUp="textLimitCheck(this, 50);"></textarea><br>限 50 个字符 已输入 <font color="#CC0000"><span id="messageCount">0</span></font> 个字                             
                            </TD>
                        </TR>
                        <TR class=tdbg>
                          <TD colspan="2" bgcolor="#FFFFFF"><div align="center">
                             <input type="hidden" name="hkuse" value="帐户充值">                         
                              <input class="inputb" type="submit" name="nextstep" value="确认订单 >>">
                            &nbsp;&nbsp;&nbsp;
                            <input class="inputb" type="reset" name="reset" value="重新填写信息">
                          </div></TD>
                        </TR>
                        </TABLE></td>
                </tr>
              </table>
                <p> </td>
          </tr>
        </form>


      </table></td>
      <td width="4" align="right" valign="top" bgcolor="#e6e6e6"></td>
    </tr>
    <tr>
      <td width="980" height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>
<SCRIPT>
	function regInput(obj, reg, inputStr)
	{
		var docSel	= document.selection.createRange()
		if (docSel.parentElement().tagName != "INPUT")	return false
		oSel = docSel.duplicate()
		oSel.text = ""
		var srcRange	= obj.createTextRange()
		oSel.setEndPoint("StartToStart", srcRange)
		var str = oSel.text + inputStr + srcRange.text.substr(oSel.text.length)
		return reg.test(str)
	}
</SCRIPT>
</body>
</html>
