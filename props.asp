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
.style1 {color: #FF0000}
-->
</style>
</head>

<body topmargin="0" leftmargin="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="371">
    <tr>
      <td width="214" valign="top">
<div align="center"><!--#include file=userleft.asp--></div></td>
<td width="5">&nbsp;</td>
      <td width="760" height="299" align="center" valign="top" bgcolor="#FFFFFF"><table width="760" height="340" border="0" align="left" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse" class="ty1">
        <tr>
          <td height="1" colspan="6"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="26">
              <tr>
                <td width="100%" height="26"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="25">
                    <tr>
                      <td width="100%" height="25" bgcolor="#F2F2F2"><span class="style1">　<strong>道具兑换<%=rmb_mc%>：</strong></span></td>
                    </tr>
                    <tr>
                      <td height="1" bgcolor="#CCCCCC"></td>
                    </tr>
                </table></td>
              </tr>
          </table></td>
        </tr>
        <%
dim jf,a,b,ct,d,hb,name,Amount
i=1
username=request.cookies("dick")("username")
set rs=server.createobject("adodb.recordset")
sql = "select jf,a,b,c,d,hb,name,Amount from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,1
jf=rs("jf")
a=trim(rs("a"))
b=trim(rs("b"))
ct=trim(rs("c"))
d=trim(rs("d"))
hb=rs("hb")
name=rs("name")
Amount=rs("Amount")
Call CloseRs(rs)
%>
        <tr>
          <td height="25" align="center" colspan="6"><p align="left">　 　此操作将你的道具兑换<%=rmb_mc%>，方便你购买其他道具！</td>
        </tr>
        <form method="POST" action="props_save.asp?dick=1&<%=C_ID%>">
          <tr>
            <td width="8%" height="20" align="center"><p align="left">　</td>
            <td width="23%" height="25" align="center"> 您有<font color="#800000">标题变色道具</font></td>
            <td width="23%" height="20" align="center"><select name="a">
                <option value="0">转换数量</option>
                <%for i=1 to int(a)%>
                <option value="<%=i%>"><%=i%>个</option>
                <%next%>
            </select></td>
            <td width="10%" height="100" rowspan="4" align="center" valign="middle"><p align="center"><img border="0" src="images/props.gif"></td>
            <td width="12%" rowspan="4" align="center" valign="middle"><%=rmb_mc%></td>
            <td width="24%" rowspan="4" align="center"><input class="inputb" type="submit" value="确定转换" name="B1"></td>
          </tr>
          <tr>
            <td width="8%" height="25" align="center"><p align="left">　</td>
            <td width="23%" height="25" align="center"> 您有<font color="#800000">信息置顶道具</font></td>
            <td width="23%" height="25" align="center"><select name="b">
                <option value="0">转换数量</option>
                <%for i=1 to int(b)%>
                <option value="<%=i%>"><%=i%>个</option>
                <%next%>
            </select></td>
          </tr>
          <tr>
            <td width="8%" height="25" align="center"><p align="left">　</td>
            <td width="23%" height="25" align="center"> 您有<font color="#800000">内容贴图道具</font></td>
            <td width="23%" height="25" align="center"><select name="ct">
                <option value="0">转换数量</option>
                <%for i=1 to ct%>
                <option value="<%=i%>"><%=i%>个</option>
                <%next%>
            </select></td>
          </tr>
          <tr>
            <td width="8%" height="25">　</td>
            <td width="23%" height="25" align="center"><p align="center">您有<font color="#800000">通过验证道具</font></td>
            <td width="23%" height="25"><p align="center">
                <select name="d">
                  <option value="0">转换数量</option>
                  <%for i=1 to d%>
                  <option value="<%=i%>"><%=i%>个</option>
                  <%next%>
                </select>
                </td>
          </tr>
        </form>

        <tr>
          <td height="25" colspan="6">　</td>
        </tr>
        <tr>
          <td height="25" colspan="6"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="25">
              <tr>
                <td width="100%" height="25" bgcolor="#F2F2F2"><span class="style1">　<strong>积分兑换<%=rmb_mc%>：</strong></span></td>
              </tr>
              <tr>
                <td height="1" bgcolor="#CCCCCC"></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td height="25" colspan="6">　　当前您的积分是：<font color="#FF0000"><%=jf%></font></td>
        </tr>
        <tr>
          <td height="25" colspan="6">　　本站可以用 <font color="#FF0000"> <%=jf_hb%></font> 积分转换 <font color="#FF0000">1</font> 个<%=rmb_mc%></td>
        </tr>
        <tr>
          <td height="25" colspan="6">　　您现有的的积分可以转换 <font color="#FF0000"><%=int(jf/jf_hb)%></font> 个<%=rmb_mc%>，剩余的积分我们将予以保留。</td>
        </tr>
        <form method="POST" action="props_save.asp?dick=2&<%=C_ID%>">
          <tr>
            <td height="25" colspan="6">　　我要转换<%=rmb_mc%>：
			<%Dim hb_max,jfi
			hb_max=int(jf/jf_hb)%>
			<select name="hb_h">
                <option value="">转换数量</option>
                <%for jfi=1 to int(hb_max)%>
                <option value="<%=jfi%>"><%=jfi%>个</option>
                <%next%>
            </select>             
              <input class="inputb" type="submit" value="确定转换" name="B1"></td>
          </tr>
        </form>

        <tr>
          <td height="25" colspan="6">　</td>
        </tr>
        <tr>
          <td height="25" colspan="6"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="25">
              <tr>
                <td width="100%" height="25" bgcolor="#F2F2F2"><span class="style1">　<strong><%=rmb_mc%>兑换积分：</strong></span></td>
              </tr>
              <tr>
                <td height="1" bgcolor="#CCCCCC"></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td height="25" colspan="6">　　当前您的<%=rmb_mc%>是：<font color="#FF0000"><%=hb%></font> 个</td>
        </tr>
        <tr>
          <td height="25" colspan="6">　　本站可以用 1个<font color="#FF0000"> <%=rmb_mc%></font> 转换 <font color="#FF0000"><%=adjfs%></font> 点积分</td>
        </tr>
        <tr>
          <td height="25" colspan="6">　　您现有的的<%=rmb_mc%>可以转换 <font color="#FF0000"><%=int(hb*adjfs)%></font> 点积分，剩余的<%=rmb_mc%>我们将予以保留。</td>
        </tr>
        <form method="POST" action="props_save.asp?dick=3&<%=C_ID%>">
          <tr>
            <td height="25" colspan="6">　　我要转换<%=rmb_mc%>：
			<%Dim jf_max
			jf_max=int(hb*adjfs/adjfs)%>
			<select name="hb_j">
                <option value="">转换数量</option>
                <%for jfi=1 to int(jf_max)%>
                <option value="<%=jfi*adjfs%>"><%=jfi*adjfs%>分</option>
                <%next%>
            </select>             
              <input class="inputb" type="submit" value="确定转换" name="B1"></td>
          </tr>
        </form>

<script LANGUAGE=JavaScript>
  function textLimitCheck2(thisArea, maxLength){
    if (thisArea.value.length > maxLength){
      alert(maxLength + ' 个字限制. \r超出的将自动去除.');
      thisArea.value = thisArea.value.substring(0, maxLength);
      thisArea.focus();
    }
    /*回写span的值，当前填写文字的数量*/
    messageCount2.innerText = thisArea.value.length;
  }
</script>
<script language="JavaScript">
function CheckForm2()
{
		if (document.form.bbb.value.length == 0) {
		alert("请输入<%=rmb_mc%>数量.");
		document.form.bbb.focus();
		return false;
	}
                if (isNaN(document.form.bbb.value))
	{
	  alert("<%=rmb_mc%>数量应为数字！")
	  document.form.bbb.focus()
	  return false
	 }

	return true;
}

</script> 
        <tr>
        <tr>
          <td height="25" colspan="6">　</td>
        </tr>
        <tr>
          <td height="25" colspan="6"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="100%">
              <tr>
                <td width="100%" height="25" bgcolor="#F2F2F2"><span class="style1">　<strong>人民币兑换<%=rmb_mc%>：</strong></span></td>
              </tr>
        <tr bgcolor="#CCCCCC">
          <td height="1"></td>
        </tr>
        <tr>
          <td height="1">          
          <center>一元人民币可兑换<font color=#ff0000><%=rmb_hb%></font>个<%=rmb_mc%>,您目前帐户中有人民币<font color=#ff0000><%=Amount%></font>元
            ,最多可兑换<font color=#ff0000><%=Amount*rmb_hb%></font>个<%=rmb_mc%>。</td>
        </tr>
        <form name="form" action="props_save.asp?dick=4&<%=C_ID%>" method="POST"  onSubmit="return CheckForm2();">
          <tr>
            <td align="center" valign="top" ><table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
                <tr>
                  <td><TABLE width="100%" height=100 border="0" align=center cellpadding="6" cellspacing="1" style="FONT-SIZE: 9pt">
                      <TBODY>
                 
                        <TR class=tdbg>
                          <TD width="16%" bgcolor="#F8F8F8">兑换<%=rmb_mc%>：</TD>
                          <TD width="86%" bgcolor="#FFFFFF"><input onKeyPress="return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,/^\d*\.?\d{0,2}$/,&#9;event.dataTransfer.getData('Text'))" maxLength=10 size=10 name="bbb" class="form">
                            个</TD>
                        </TR>

                        <TR class=tdbg>
                          <TD width="16%" bgcolor="#F8F8F8">备&nbsp;&nbsp;&nbsp;&nbsp;注：</TD>
                          <TD bgcolor="#FFFFFF">
      <textarea name="czbz" cols="50" rows="3" onkeyUp="textLimitCheck2(this, 50);">人民币兑换<%=rmb_mc%></textarea><br>限 50 个字符 已输入 <font color="#CC0000"><span id="messageCount2">0</span></font> 个字                             
                            </TD>
                        </TR>	
				 <input type="hidden" name="aliname" value="<%=name%>">				 
				 <input type="hidden" name="Amount2" value="<%=Amount%>">
				 <input type="hidden" name="rmb_hb" value="<%=rmb_hb%>">				
                        <TR class=tdbg>
                          <TD colspan="2" bgcolor="#FFFFFF"><div align="center">
                              <input type="hidden" name="hkuse" value="升级VIP">
                              <input class="inputb" type="submit" name="nextstep" value="确认转换 >>">
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
	
		if (document.alipayment.Amount2.value.length == 0) {
		alert("请输入人民币数量.");
		document.alipayment.Amount2.focus();
		return false;
	}
                if (isNaN(document.alipayment.Amount2.value))
	{
	  alert("人民币数量应为数字！")
	  document.alipayment.Amount2.focus()
	  return false
	 }

	return true;

}

</script>
        <tr>
        <tr>
          <td height="25" colspan="6">　</td>
        </tr>
        <tr>
          <td height="25" colspan="6"><!--table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="25">
              <tr>
                <td width="100%" height="25" bgcolor="#F2F2F2"><span class="style1">　<strong><%=rmb_mc%>兑换人民币：</strong></span></td>
              </tr>
              <tr>
                <td height="1" bgcolor="#CCCCCC"></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td height="25" colspan="6"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="100%">

        <tr bgcolor="#CCCCCC">
          <td height="1"></td>
        </tr>
        <tr>
          <td height="1">          
          <center>一个<%=rmb_mc%>可兑换<font color=#ff0000><%=rmb_hb%></font>元人民币,您目前帐户中有<%=rmb_mc%><font color=#ff0000><%=hb%></font>个,最多可兑换<font color=#ff0000><%=rmb_hb*hb%></font>元人民币。</td>
        </tr>
        <form name="alipayment" action="props_save.asp?dick=5&<%=C_ID%>" method="POST"  onSubmit="return CheckForm();">
          <tr>
            <td align="center" valign="top" ><table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
                <tr>
                  <td><TABLE width="100%" height=100 border="0" align=center cellpadding="6" cellspacing="1" style="FONT-SIZE: 9pt">
                      <TBODY>

                        <TR class=tdbg>
                          <TD width="16%" bgcolor="#F8F8F8">兑换人民币：</TD>
                          <TD width="86%" bgcolor="#FFFFFF"><input onKeyPress="return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,/^\d*\.?\d{0,2}$/,&#9;event.dataTransfer.getData('Text'))" maxLength=10 size=10 value="" name="Amount2" class="form"> 元</TD>
                        </TR>
                        <TR class=tdbg>
                          <TD width="16%" bgcolor="#F8F8F8">备&nbsp;&nbsp;&nbsp;&nbsp;注：</TD>
                          <TD bgcolor="#FFFFFF">
      <textarea name="czbz" cols="50" rows="3" onkeyUp="textLimitCheck(this, 50);"><%=rmb_mc%>转人民币</textarea><br>限 50 个字符 已输入 <font color="#CC0000"><span id="messageCount">0</span></font> 个字                             
                            </TD>
                        </TR>				 
				 <input type="hidden" name="aliname" value="<%=name%>">	
				 <input type="hidden" name="hb" value="<%=hb%>">
                 <input type="hidden" name="rmb_hb" value="<%=rmb_hb%>">				
                        <TR class=tdbg>
                          <TD colspan="2" bgcolor="#FFFFFF"><div align="center">
                              <input type="hidden" name="hkuse" value="升级VIP">
                              <input class="inputb" type="submit" name="nextstep" value="确认转换 >>">
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


      </table---></td>
          </tr>


      </table></td>
          </tr>
		  
        <tr>
          <td height="25" colspan="6">　</td>
        </tr>
      </table></td>
      <td width="4" align="center" valign="top" bgcolor="#e6e6e6"></td>
    </tr>
    <tr>
      <td height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>
</body>
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
</html>
