<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->
<!--#include file=usercookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
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
      alert(maxLength + ' ��������. \r�����Ľ��Զ�ȥ��.');
      thisArea.value = thisArea.value.substring(0, maxLength);
      thisArea.focus();
    }
    /*��дspan��ֵ����ǰ��д���ֵ�����*/
    messageCount.innerText = thisArea.value.length;
  }
</script>
<script language="JavaScript">
function CheckForm()
{
	
	
		if (document.alipayment.vipys.value.length == 0) {
		alert("������VIP��������.");
		document.alipayment.vipys.focus();
		return false;
	}
                if (isNaN(document.alipayment.vipys.value))
	{
	  alert("VIP��������ӦΪ���֣�")
	  document.alipayment.vipys.focus()
	  return false
	 }

	return true;
}

</script>
</head>

<body topmargin="0" leftmargin="0">
<%
dim name,Amount
%>

<div align="center">
  <center>
  <table width="760" height="423" border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" bgcolor="#FFFFFF" style="border-collapse: collapse">
    <tr>
      <td width="214" height="351" valign="top"><div align="center">
            <!--#include file=userleft.asp-->��
        </div></td>
<td width="5">&nbsp;</td>        
<%Dim ywlx
username=request.cookies("dick")("username")
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_user] where username='"&username&"'" 
rs.open sql,conn,1,1
if rs.eof then
response.write "<br>"
response.write "<li>��������"
response.end
end if
name=rs("name")
Amount=rs("Amount")
ywlx="֧��"
vip=rs("vip")
Call CloseRs(rs)%>
      <td width="760" align="right" valign="top" bgcolor="#FFFFFF"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="100%" class="ty1">
        <tr bgcolor="#FAFAFA">
          <td height="25"><span class="style5">������VIP��Ա��</span></td>
        </tr>
        <tr bgcolor="#CCCCCC">
          <td height="1"></td>
        </tr>
        <tr>
          <td height="1">          
          <center>����ΪVIP��Ա�������<font color=#ff0000><%=vipje%></font>Ԫ/��,��Ŀǰ�ʻ����������<font color=#ff0000><%=Amount%></font>Ԫ
            ,���ɹ���VIP�ʸ�<font color=#ff0000><%=Amount/vipje%></font>���£�<b>��վҪ������һ���Թ���<font color=#ff0000><%=vipsj%></font>���£�</b>��</td>
        </tr>
        <form name="alipayment" action="upgradevip_save.asp?<%=C_ID%>" method="POST"  onSubmit="return CheckForm();">
          <tr>
            <td align="center" valign="top" ><table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
                <tr>
                  <td><TABLE width="100%" height=100 border="0" align=center cellpadding="6" cellspacing="1" style="FONT-SIZE: 9pt">
                      <TBODY>
                   <TR class=tdbg>
                          <TD width="16%" bgcolor="#F8F8F8">�����ʺţ�</TD>
                          <TD width="86%" bgcolor="#FFFFFF"><b><font color="#0000FF"><%=username%></font></b>
						  ��<%If vip=1 Then
                          response.write "VIP��Ա"
						  Else
						  response.write "��ͨ��Ա"
						  End if%>��</TD>
                        <TR class=tdbg>
                          <TD width="16%" bgcolor="#F8F8F8">����������</TD>
                          <TD width="86%" bgcolor="#FFFFFF"><input onKeyPress="return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,/^\d*\.?\d{0,2}$/,&#9;event.dataTransfer.getData('Text'))" maxLength=10 size=10 name="vipys" value="<%=vipsj%>" class="form">
                            ���£�ָ����VIP��Ա�ʸ������,���¼ƣ�ÿ����30�죬VIP��Ա�������<font color=#ff0000><%=vipje%></font>Ԫ/�£�</TD>
                        </TR>

                        <TR class=tdbg>
                          <TD width="16%" bgcolor="#F8F8F8">��&nbsp;&nbsp;&nbsp;&nbsp;ע��</TD>
                          <TD bgcolor="#FFFFFF">
      <textarea name="czbz" cols="50" rows="3" onkeyUp="textLimitCheck(this, 50);">����VIP��Ա�ʸ�</textarea><br>�� 50 ���ַ� ������ <font color="#CC0000"><span id="messageCount">0</span></font> ����                             
                            </TD>
                        </TR>
	
				 <input type="hidden" name="username" value="<%=username%>">
				 <input type="hidden" name="aliname" value="<%=name%>">
				 <input type="hidden" name="ywlx" value="<%=ywlx%>">
				 <input type="hidden" name="Amount" value="<%=Amount%>">
				 <input type="hidden" name="vipje" value="<%=vipje%>">
				  <input type="hidden" name="vip" value="<%=vip%>">
                        <TR class=tdbg>
                          <TD colspan="2" bgcolor="#FFFFFF"><div align="center">
                              <input type="hidden" name="hkuse" value="����VIP">
                              <input class="inputb" type="submit" name="nextstep" value="ȷ������ >>">
                            &nbsp;&nbsp;&nbsp;
                            <input class="inputb" type="reset" name="reset" value="������д��Ϣ">
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
      <td width="760" height="26" colspan="3"><!--#include file=footer.asp--></td>
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
