<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
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
//�ֽ������ƣ����������=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}
//��ʾ�����ַ���

function kn()
  {
   var ln=thisForm.username.value.Length();
     num.innerText='������������:'+(16-ln)+'���ַ�';
      //if (ln>=16) form.username.readOnly=true;  //��������������������޷�������޸�
}
function kn2()
 {
   var ln=thisForm.password.value.Length();
     num2.innerText='������������:'+(12-ln)+'���ַ�';
      
}
// --></script>
<script language="javascript" type="text/javascript">
<!--
//��֤������ȷ��
function checkemail(email){
var str=email;
//��JavaScript�У�������ʽֻ��ʹ��"/"��ͷ�ͽ���������ʹ��˫����
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
	    alert("��¼�ʺŲ���Ϊ��!");
		document.thisForm.username.focus()
	    return false;
	}
     if ((document.thisForm.username.value.Length()>16) || (document.thisForm.username.value.Length()<1)) //�ֽ������ƣ����������
     {
	 alert("��¼�ʺų���Ҫ��1��16�ֽڣ����������룡");
	  document.thisForm.username.focus()
	  return false
  }

	  if ((document.thisForm.password.value.Length()>12) || (document.thisForm.password.value.Length()<5)) //�ֽ������ƣ����������
     {
	 alert("���볤��Ҫ��5��12�ֽڣ����������룡");
	  document.thisForm.password.focus()
	  return false
  }
	    if(document.thisForm.password1.value.length<1)
	{
	    alert("ȷ�����벻��Ϊ��!");
		document.thisForm.password1.focus()
	    return false;
	}
	   if(document.thisForm.password1.value!=document.thisForm.password.value)
        {
            alert("�����������벻һ��!");
			document.thisForm.password1.focus()
            return false;
        }
	if((!checkemail(thisForm.email.value))&&(document.thisForm.email.value!=""))
    {
    alert("������Email��ַ����ȷ������������!");
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
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">ֱ �� �� �� �� ��</FONT></TD>
 </TR>
  <tr bgcolor="#FFFFFF">
    <td height="44" style="border-top-style: solid; border-top-width: 1"><table width="100%"  border="0" cellspacing="6" cellpadding="6">
      <tr>
        <td width="16%">�� �� ����</td>
        <td width="84%"><input type="text" name="username" maxlength="16" size="30"> <font color="#ff0000"> *</font></td>
      </tr>
      <tr>
        <td width="16%">��&nbsp;&nbsp;&nbsp;&nbsp;�룺</td>
        <td width="84%"><input type="text" name="password" size="30"> <font color="#ff0000"> *</font></td>
      </tr>
      <tr>
        <td width="16%">ȷ�����룺</td>
        <td width="84%"><input type="text" name="password1" size="30"> <font color="#ff0000"> *</font></td>
      </tr>
      <tr>
        <td>�ʼ���ַ��</td>
        <td><input type="text" name="email" size="30"> ������д����</td>
      </tr>
            <tr>
        <td>������Ѷ��</td>
        <td><input type="radio" name="maillist" value="1" >                 
                  ����&nbsp;&nbsp;&nbsp;&nbsp;                          
                  <input type="radio" name="maillist" value="0" checked>                 
                  ������</td>
      </tr>
    </table></td>
    </tr>
  <tr bgcolor="#eeeeee">
    <td height="35"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><div align="center">
            <input type="submit" value="ȷ�����" name="B1">
        </div></td>
        <td><div align="center">
            <input type="reset" value="ȡ�����" name="B1">
        </div></td>
      </tr>
    </table></td>
    </tr>
  </form>
</table>
  </center>
</div>
