<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<%call checkmanage("14")%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script LANGUAGE="JavaScript">
function FormCheck(theform)
{
if (document.form1.wenti.value=="")
{
alert("���������⣡")
document.form1.wenti.focus()
document.form1.wenti.select()
return false; 
}
if (document.form1.daan.value=="")
{
alert("������𰸣�")
document.form1.daan.focus()
document.form1.daan.select() 
return false; 
}
<!-- theform.submit() -->
}
</script>
<title>�������ʹ�</title></head>
<body bgcolor="#FFFFFF" marginheight=0 marginwidth=0 leftmargin=0 background="../images/back.gif">
<link rel="stylesheet" href="../css/css.css" type="text/css">
<form name="form1" method="POST" action="admin_check_addsave.asp" onsubmit="return FormCheck(this);">
<div align="center">
<center>
<TABLE border="0" cellspacing="1" width="570" style="border-collapse: collapse">
<TR>
<TD width="601" height="20" align="center"><b><font color="#ff6600" class="f12">
�������ʹ�</font></b></TD>
</TR>
<TR ALIGN="center">
<TD width="601">
<TABLE border="1" cellspacing="1" width="574" bordercolorlight="#CEE7FF" bordercolordark="#CEE7FF" style="border-collapse: collapse" bordercolor="#111111" height="107">
<TR> 
<TD width="574" align="right" bgcolor="#F4FAFF" height="26">���⣺</TD>
<TD width="888" height="26" bgcolor="#F4FAFF">
<font color="#F4FAFF"><input name="wenti" type="text" class="Input" size="50">
</font> *</TD>
</TR>
<TR> 
<TD width="574" align="right" bgcolor="#F4FAFF" rowspan="2" height="130">�𰸣�</TD>
<TD width="888" bgcolor="#F4FAFF" height="27">
<font color="#F4FAFF">
<input name="daan" type="text" class="input" size="50">
</font> *</TD>
</TR>
<tr>
<TD width="888" bgcolor="#F4FAFF" height="17">��</TD>
</tr>
<TR height="40">
<TD colspan="2" align="center" width="574" bgcolor="#F4FAFF" height="28">
<font color="#F4FAFF"><input type="submit" name="Submit" value=" �ᡡ��" class="Input">
</font>
<font color="#F4FAFF">
<input type="reset" name="Submit2" value=" �ء�ִ" class="Input">
</font>
</TR>
</TABLE>
</TD>
</TR>
</TABLE>
</center>
</div>
</FORM>
</BODY>
</HTML>


