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
<%Dim problem
Call OpenConn
problem=trim(conn.Execute("Select problem From DNJY_user Where username='"&request.cookies("dick")("username")&"'")(0))
Call CloseDB(conn)%>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<link rel="stylesheet" type="text/css" href="../1.CSS">
<style type="text/css">
<!--
.style1 {color: #FF0000}
.style5 {color: #FF0000; font-weight: bold; }
-->
</style>
<script language="javascript">
<!--
//�ֽ������ƣ����������=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}
// --></script>
<script language="javascript" type="text/javascript">
<!--
function CheckForm()
{

        if(document.thisForm.pass1.value.length<1)
	{
	    alert("ԭ���벻��Ϊ��!");
		document.thisForm.pass1.focus()
	    return false;
	}
	  if ((document.thisForm.pass2.value.Length()>12) || (document.thisForm.pass2.value.Length()<5)) //�ֽ������ƣ����������
     {
	 alert("�����볤��Ҫ��5��12�ֽڣ����������룡");
	  document.thisForm.pass2.focus()
	  return false
  }
	    if(document.thisForm.pass3.value.length<1)
	{
	    alert("ȷ�������벻��Ϊ��!");
		document.thisForm.pass3.focus()
	    return false;
	}
	   if(document.thisForm.pass3.value!=document.thisForm.pass2.value)
        {
            alert("�����������벻һ��!");
			document.thisForm.pass3.focus()
            return false;
        }
}


  function CheckForm_answer()
    {
<%if IsNull(problem) then%>
        if(document.Form.problem.value.length<1)
	{
	    alert("�������ⲻ��Ϊ��!");
		document.Form.problem.focus()
	    return false;
	}
        if(document.Form.answer1.value.length<1)
	{
	    alert("�𰸲���Ϊ��!");
		document.Form.answer1.focus()
	    return false;
	}
	    if(document.Form.answer2.value.length<1)
	{
	    alert("ȷ�ϴ𰸲���Ϊ��!");
		document.Form.answer2.focus()
	    return false;
	}
	   if(document.Form.answer1.value!=document.Form.answer2.value)
        {
            alert("��������𰸲�һ��!");
			document.Form.answer1.focus()
            return false;
        }
<%else%>
        if(document.Form.problem.value.length<1)
	{
	    alert("���������ⲻ��Ϊ��!");
		document.Form.problem.focus()
	    return false;
	}
        if(document.Form.answer1.value.length<1)
	{
	    alert("ԭ�𰸲���Ϊ��!");
		document.Form.answer1.focus()
	    return false;
	}
	    if(document.Form.answer2.value.length<1)
	{
	    alert("ȷ���´𰸲���Ϊ��!");
		document.Form.answer2.focus()
	    return false;
	}
	   if(document.Form.answer3.value!=document.Form.answer2.value)
        {
            alert("��������𰸲�һ��!");
			document.Form.answer3.focus()
            return false;
        }
<%end if%>

}
//-->
</SCRIPT>
</head>

<body background="img/bg1.gif" leftmargin="0" topmargin="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="426">
    <tr>
      <td width="214" height="354" valign="top" bgcolor="#FFFFFF"><div align="center">
      
        <!--#include file=userleft.asp--></div></td>
		<td width="5">&nbsp;</td>
      <td width="760" height="354" valign="top" bgcolor="#FFFFFF">        <div align="right"></div>
        <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse" class="ty1">
          
            <tr>
              <td height="25" bgcolor="#FAFAFA"><span class="style5">�û���������뱣���ʴ��޸ģ�</span></td>
              <td bgcolor="#FAFAFA">&nbsp;</td>
            </tr>
            <tr bgcolor="#CCCCCC">
              <td width="456" height="1"></td>
              <td width="17"></td>
            </tr>
            <tr>
              <td width="500" valign="top" align="center">
			  <form name="thisForm"  method="POST" action="user_pass_save.asp?key=pass"  language="javascript" onsubmit="return CheckForm()">
			  <table border="0" cellspacing="0" cellpadding="0" height="212">
                  <tr>
                    <td height="22" colspan="3"><p align="center"><b>�޸��û���¼����</b></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right">����ԭ���룺</td>
                    <td  height="30">&nbsp;
                        <input class="inputa"��type="text" name="pass1" size="20" maxlength="12"></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right">���������룺</td>
                    <td height="30">&nbsp;
                        <input class="inputa" type="password" name="pass2" size="20"></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right"> ȷ�������룺</td>
                    <td height="30">&nbsp;
                        <input class="inputa" type="password" name="pass3" size="20"></td>
                  </tr>

                  <tr>
                    <td height="25" colspan="3" align="center"><table  width="200" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td><div align="center">
                              <input name="B1" type="submit" class="inputb" value="�ύ�޸�" border="0"   onclick="javascript:return CheckForm();">
                          </div></td>
                          <td><div align="center">
                              <input name="B2" type="reset" class="inputb" id="B2" value="ȡ���޸�" border="0" >
                          </div></td>
                        </tr>
                    </table></td>
                    <td height="25" width="66"> ���������� </td>
                  </tr>
                  <tr>
                    <td height="20" colspan="3"><div align="center">����д��ȷ�ľ����룬����������뽫<span class="style1">�޷��޸�</span>��</div></td>
                  </tr>

              </table>
			 </form>
<TABLE>
<TR>
	<TD height="22"></TD>
</TR>
</TABLE>
			 <form name="Form"  method="POST" action="user_pass_save.asp?key=answer"  language="javascript" onsubmit="return CheckForm_answer()">
                 <table border="0" cellspacing="0" cellpadding="0" height="212">

				  <%If IsNull(problem) then%>
                  <tr>
                    <td height="22" colspan="3"><p align="center"><b>����δ�������뱣�����뼰ʱ����</b></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right">�����ܱ����⣺</td>
                    <td  height="30">&nbsp;
                        <input class="inputa"��type="text" name="problem" size="20" maxlength="12"></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right">�����ܱ��𰸣�</td>
                    <td height="30">&nbsp;
                        <input class="inputa" type="text" name="answer1" size="20"></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right"> ȷ���ܱ��𰸣�</td>
                    <td height="30">&nbsp;
                        <input class="inputa" type="text" name="answer2" size="20"></td>
                  </tr>
				  <%else%>
                  <tr>
                    <td height="22" colspan="3"><p align="center"><b>�޸����뱣������</b></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right">�ܱ����⣺</td>
                    <td  height="30">&nbsp;<%=problem%></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right">���������⣺</td>
                    <td  height="30">&nbsp;
                        <input class="inputa"��type="text" name="problem" size="20" maxlength="12"></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right">����ԭ�𰸣�</td>
                    <td  height="30">&nbsp;
                        <input class="inputa"��type="text" name="answer1" size="20" maxlength="12"></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right">�����´𰸣�</td>
                    <td height="30">&nbsp;
                        <input class="inputa" type="text" name="answer2" size="20"></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right"> ȷ���´𰸣�</td>
                    <td height="30">&nbsp;
                        <input class="inputa" type="text" name="answer3" size="20"></td>
                  </tr>
				  <input type="hidden" name="Modify" value="ok">
                <%End if%>
                  <tr>
                    <td height="25" colspan="3" align="center"><table  width="200" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td><div align="center">
                              <input name="B1" type="submit" class="inputb" value="ȷ��" border="0"   onclick="javascript:return CheckForm_answer();">
                          </div></td>
                          <td><div align="center">
                              <input name="B2" type="reset" class="inputb" id="B2" value="ȡ��" border="0" >
                          </div></td>
                        </tr>
                    </table></td>
                    <td height="25" width="66"> ���������� </td>
                  </tr>
              </table>
           </form>
			  </td>
              <td width="17"> ��</td>
            </tr>
         
        </table></td>
      <td width="4" valign="top" bgcolor="#e6e6e6"></td>
    </tr>
    <tr>
      <td height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>