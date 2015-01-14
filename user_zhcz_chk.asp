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
<%Response.Buffer = True    
Response.ExpiresAbsolute = Now() - 1    
Response.Expires = 0    
Response.CacheControl = "no-cache"

dim aliname,alimoney,dianhua,CompPhone,email,oicq,alibody,dizhi,youbian,hkuse
%>
<%
aliname=request.form("aliname")
alimoney=request.form("alimoney")
dianhua=request.form("dianhua")
CompPhone=request.form("CompPhone")
email=request.form("email")
oicq=request.form("oicq")
dizhi=request.form("dizhi")
youbian=request.form("youbian")
alibody=request.form("alibody")
hkuse=request.form("hkuse")


%>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<style type="text/css">
<!--
.style5 {color: #FF0000; font-weight: bold; }
.style6 {color: #FF0000}
-->
</style>
<script language="JavaScript">
function CheckForm()
{

	if (document.alipayment.aliname.value.length == 0) {
		alert("请输入购买人的姓名.");
		document.alipayment.aliname.focus();
		return false;
	}
         if (isNaN(document.alipayment.alimoney.value))
	{
	  alert("付款金额应为数字！")
	  document.alipayment.alimoney.focus()
	  return false
	 }
	if (document.alipayment.dianhua.value.length == 0) {
		alert("请购买人的电话.");
		document.alipayment.dianhua.focus();
		return false;
	}

	return true;
}

</script>
</head>

<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
  <table width="760" height="423" border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" bgcolor="#FFFFFF" style="border-collapse: collapse">
    <tr>
      <td width="214" height="351" valign="top"><div align="center">
        <!--#include file=userleft.asp--></div></td>
		<td width="5">&nbsp;</td>
      <td width="760" height="351" valign="top" bgcolor="#FFFFFF"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="100%" class="ty1">
        <tr bgcolor="#FAFAFA">
          <td height="25"><span class="style5">　购买项目：<%=hkuse%></span></td>
        </tr>
        <tr bgcolor="#CCCCCC">
          <td height="1"></td>
        </tr>
        <tr>
          <td height="1"> <center>请检查所填信息是否正确<br>
            </td>
        </tr>
        <tr>
          <td align="center"><table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
              <tr>
<%'防止页面刷新重复提交,与INPUT和下一表单处理页相结合
dim num1
Randomize '初始代随机数种子
num1=rnd() '产生随机数num1
num1=int(26*num1)+65 '修改num1的范围以使其是A-Z范围的Ascii码，以防表单名出错
session("antry")="test"&chr(num1) '产生随机字符串
%>

                <td><form name="alipayment" method="post" action="user_zhcz_save.asp?<%=C_ID%>">
                    <TABLE width="100%" height=137 border="0" align=center cellpadding="6" cellspacing="1" style="FONT-SIZE: 9pt">
                      <TBODY>

                        <TR class=tdbg>
                          <TD width="16%" bgcolor="#F8F8F8">您的帐号：</TD>
                          <TD width="86%" bgcolor="#FFFFFF"><b><font color="#0000FF"><%=request.cookies("dick")("username")%></font></b></TD>
                        </TR>                        
                        <TR class=tdbg>
                          <TD width="16%" bgcolor="#F8F8F8">购 买 人：</TD>
                          <TD width="86%" bgcolor="#FFFFFF"><%=aliname%></TD>
                        </TR>
                        <TR class=tdbg>
                          <TD width="16%" bgcolor="#F8F8F8">付款金额：</TD>
                          <TD width="86%" bgcolor="#FFFFFF"><%=alimoney%>
                            元人民币</span></TD>
                        <TR class=tdbg>
                          <TD bgcolor="#F8F8F8">购买项目：</TD>
                          <TD bgcolor="#FFFFFF"><%=hkuse%></TD>
                        </TR>                            
                        <TR class=tdbg>
                          <TD bgcolor="#F8F8F8">固定电话：</TD>
                          <TD bgcolor="#FFFFFF"><%=dianhua%></TD>
                        </TR>
                        <TR class=tdbg>
                          <TD bgcolor="#F8F8F8">移动电稿：</TD>
                          <TD bgcolor="#FFFFFF"><%=CompPhone%></TD>
                        </TR>
                        <TR class=tdbg>
                          <TD bgcolor="#F8F8F8">电子信箱：</TD>
                          <TD bgcolor="#FFFFFF"><%=email%></TD>
                        </TR>                        
                        <TR class=tdbg>
                          <TD bgcolor="#F8F8F8">腾讯OICQ：</TD>
                          <TD bgcolor="#FFFFFF"><%=oicq%></TD>
                        </TR>
                        <TR class=tdbg>
                          <TD bgcolor="#F8F8F8">腾讯地址：</TD>
                          <TD bgcolor="#FFFFFF"><%=dizhi%></TD>
                        </TR>
                        <TR class=tdbg>
                          <TD bgcolor="#F8F8F8">邮政编码：</TD>
                          <TD bgcolor="#FFFFFF"><%=youbian%></TD>
                        </TR>                                               

                        <TR class=tdbg>
                          <TD bgcolor="#F8F8F8">客户留言：</TD>
                          <TD bgcolor="#FFFFFF" style="word-break:break-all"><%=alibody%></TD>
                        </TR>
<!--下面一行为防止页面刷新重复提交，与form前代码配合--->                        
<input type="hidden" name='<%=session("antry")%>' value='<%=session("antry")%>'>

<INPUT TYPE="hidden" name="aliname" value="<%=aliname%>">
<INPUT TYPE="hidden" name="alimoney" value="<%=alimoney%>">
<INPUT TYPE="hidden" name="dianhua" value="<%=dianhua%>">
<INPUT TYPE="hidden" name="CompPhone" value="<%=CompPhone%>">
<INPUT TYPE="hidden" name="hkuse" value="<%=hkuse%>">
<INPUT TYPE="hidden" name="email" value="<%=email%>">
<INPUT TYPE="hidden" name="oicq" value="<%=oicq%>">
<INPUT TYPE="hidden" name="dizhi" value="<%=dizhi%>">
<INPUT TYPE="hidden" name="youbian" value="<%=youbian%>">
<INPUT TYPE="hidden" name="alibody" value="<%=alibody%>">

                        <TR class=tdbg>
                          <TD colspan="2" bgcolor="#FFFFFF"><div align="center">
                              <input class="inputb" type="submit" name="Submit" value="提交充值订单">
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                          <input class="inputb" type="button" name="Submit21" onClick="javascript:history.go(-1)" value="返回修改订单">  
                          </div></TD>
                        </TR>
                      </TBODY>
                    </TABLE>
                </form></td>
              </tr>
            </table>
              <p> </td>
        </tr>


      </table></td>
      <td width="4" align="right" valign="top" bgcolor="#e6e6e6"></td>
    </tr>
    <tr>
      <td width="760" height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>
