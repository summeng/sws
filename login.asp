<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<style>
<!--
.a1 {
	color: #FFFF00;
	text-decoration: none;
}
.style1 {color: #FF0000}
-->
</style>
</head>
<SCRIPT language=javascript>
<!--
function CheckForm()
{
if(document.thisForm.username2.value.length<1)
	{
	    alert("�û�������Ϊ��!");
	    return false;
	}
	if(document.thisForm.password.value.length<1)
	{
	    alert("���벻��Ϊ��!");
	    return false;
	}
	if(document.thisForm.yzm.value.length<1)
	{
	    alert("��֤�벻��Ϊ��!");
	    return false;
	}
}

//-->
</SCRIPT>
<body topmargin="0" leftmargin="0">

<div align="center">
  <center>
    <table width="980" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
      <tr>
        <td height="4" colspan="3" bgcolor="#e6e6e6"></td>
      </tr>
      <tr>
        <td width="4" bgcolor="#e6e6e6"></td>
        <td><table width="700" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
          <tr>
            <td height="1" colspan="2"></td>
          </tr>
          <tr>
            <td height="30">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td><img src="img/login.jpg" width="171" height="56"></td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td width="400"><table width="400" border="0" align="center" cellPadding="0" cellSpacing="0" borderColor="#111111" class="font_10_e_blue" style="BORDER-COLLAPSE: collapse">

                  <tr>
                    <td height="69" colspan="2" vAlign="top"><table class="font_10_e_black" cellSpacing="0" cellPadding="3" width="100%" align="center" border="0">
                        <form id="f1" name="thisForm" action="login_save.asp?<%=C_ID%>" method="post">
                          
                          <tr>
                            <td align="right">��½�ʺţ�</td>
                            <td><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" class="form_e_10_black" id="username2" maxLength="20" size="25" name="username2"></td>
                          </tr>
                         <tr>
                            <td align="right">��½���룺</td>
                            <td><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="password" maxLength="32" size="25" value name="password"></td>
                          </tr>
                          <tr>
                            <td align="right">�� ֤ �룺</td>
                            <td><input type="text" name="yzm" size="4" /><img src="tt_getcode.asp" alt="��֤��,�������?����ˢ����֤��" height="18" style="cursor : pointer;" onClick="this.src='tt_getcode.asp?t='+(new Date().getTime());" /> <input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;width:48px" onClick="javascript:return CheckForm();" type="submit" value="��½"  border="0" name="I2"></TD>
                          </tr>
                        </form>
                    </table><br></td>
                  </tr>
<TR><%Dim qq_login
If API_QQEnable=True And (API_SinaEnable=false or API_AlipayEnable=false) Then
qq_login="qq_login"
Else
qq_login="qqSmall"
End if%>
	<TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%If API_QQEnable=True then%><%response.write " <a title=""ʹ��QQ�ſ��ٵ�¼"" href=""" & strInstallDir & "api/qq/redirect_to_login.asp""><img align=""absmiddle"" src=""" & strInstallDir &"api/img/"&qq_login&".png"" border=""0"" align=""absmiddle""/></a>"%><%End if%><%If API_SinaEnable=True then%>&nbsp;&nbsp;<%response.write " <a title=""ʹ������΢���ʺŵ�¼"" href=""" & strInstallDir & "api/sina/deal.asp""><img align=""absmiddle"" src=""" & strInstallDir &"api/img/sinaSmall.png"" border=""0"" align=""absmiddle""/></a>"%><%End if%><%If API_AlipayEnable=True then%>&nbsp;&nbsp;<%response.write " <a title=""ʹ��֧�����ʺŵ�¼"" href=""" & strInstallDir & "api/alipay/alipay_auth_authorize.asp""><img align=""absmiddle"" src=""" & strInstallDir &"api/img/alipay_button.gif"" border=""0"" align=""absmiddle""/></a>"%><%End if%></TD>
</TR>
                  <tr>
                    <td><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<BUTTON 
                        onclick="window.open('user/RetakePassword_a.asp?<%=C_ID%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=420,height=175,left=300,top=300');" 
                        style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: 1px solid #a2a2a2; BORDER-LEFT: 1px solid #ffffff; BORDER-RIGHT: 1px solid #a2a2a2; BORDER-TOP: 1px solid #ffffff; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px; width:70px">
                       <div align="center">�����һ�</div></BUTTON>&nbsp;&nbsp;&nbsp;&nbsp;<BUTTON 
                        onclick="window.open('<%=reg%>?<%=C_ID%>');" 
                        style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: 1px solid #a2a2a2; BORDER-LEFT: 1px solid #ffffff; BORDER-RIGHT: 1px solid #a2a2a2; BORDER-TOP: 1px solid #ffffff; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px; width:70px">
                              <div align="center">���ע��</div></BUTTON>&nbsp;&nbsp;&nbsp;&nbsp;<BUTTON 
                        onclick="window.open('help.asp?id=1&<%=C_ID%>');" 
                        style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: 1px solid #a2a2a2; BORDER-LEFT: 1px solid #ffffff; BORDER-RIGHT: 1px solid #a2a2a2; BORDER-TOP: 1px solid #ffffff; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px; width:70px">
                     <div align="center">��������</div>
                     </BUTTON></td>
                  </tr>
                  <tr>
                    <td width="173" height="17" colspan="2" vAlign="top" bgcolor="#FFFFFF"> <div align="center"></div></td>
                  </tr>
                </table>
           ��            
              </td>
            <td width="400"><table width="372" border="0" align="center" cellPadding="0" cellSpacing="0" bordercolor="#111111" style="border-collapse: collapse">
               <tr vAlign="top">
                  <td align="center" height="25" width="44"><img src="images/form2_r3_c3.gif"></td>
                  <td align="left" height="25" width="328">�������¼��������²�����</td>
                </tr>
                <tr vAlign="top">
                  <td align="center" height="25" width="44"><img src="images/form2_r3_c3.gif"></td>
                  <td align="left" height="25" width="328">��ֻ��<a class="a1" href="<%=reg%>?<%=C_ID%>"><font color="#FF6600">���ע��</font></a>һ�Σ��Ϳ���������ͨ��Ա�����������ΪVIP��Ա��������ܸ�����Ȩ��</td>
                </tr>
                <tr vAlign="top">
                  <td align="center" height="25" width="44"><img src="images/form2_r3_c3.gif"></td>
                  <td align="left" height="25" width="328">�������������룬<a href=# onClick="javascript:window.open('user/RetakePassword_a.asp?<%=C_ID%>','','scrollbars=1,width=420,height=175,left=300,top=300,toolbar=0,resizable=0');"><font color="#FF6600">�ɵ���һ�����</font></a>��</td>
                </tr>
                <tr>
                  <td vAlign="top" align="center" height="25" width="44"><img src="images/form2_r3_c3.gif"></td>
                  <td vAlign="top" align="left" height="25" width="328"> ��û���ر�˵��ʱ�����ǵķ�����<font color="#FF6600">��ѵ�</font>�������ʹ�á�</td>
                </tr>
                <tr>
                  <td vAlign="top" align="center" height="27" width="44"><img src="images/form2_r3_c3.gif"></td>
                  <td vAlign="top" align="left" height="27" width="328">VIP������ʹ��<font color="#FF6600"><%=rmb_mc%></font>��Ϊ֧���ֶεģ�����<font color="#FF6600"><%=rmb_mc%></font>��������ڱ�վ�κεط��������ѡ�</td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="26"></td>
            <td height="26"></td>
          </tr>
        </table></td>
        <td width="4" bgcolor="#e6e6e6"></td>
      </tr>
    </table>
  </center>
</div>

</body>

</html>
<!--#include file=footer.asp-->