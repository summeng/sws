<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../Include/md5.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("01")%>
<script language=javascript src=../Include/mouse_on_title.js></script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<style type="text/css">
<!--
.style1 {
	color: #FFFFFF;
	font-weight: bold;
}
body {
	background-color: #E3EBF9;
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
// --></script>
<script language="javascript" type="text/javascript">
<!--
function CheckForm()
{

        if(document.thisForm.Pass0.value.length<1)
	{
	    alert("ԭ���벻��Ϊ��!");
		document.thisForm.Pass0.focus()
	    return false;
	}
	  if ((document.thisForm.Pass1.value.Length()>16) || (document.thisForm.Pass1.value.Length()<8)) //�ֽ������ƣ����������
     {
	 alert("�����볤��Ҫ��8��16�ֽڣ����������룡");
	  document.thisForm.Pass1.focus()
	  return false
  }
	    if(document.thisForm.Pass2.value.length<1)
	{
	    alert("ȷ�������벻��Ϊ��!");
		document.thisForm.Pass2.focus()
	    return false;
	}
	   if(document.thisForm.Pass2.value!=document.thisForm.Pass1.value)
        {
            alert("�����������벻һ��!");
			document.thisForm.Pass2.focus()
            return false;
        }

  
}
//-->
</SCRIPT>
</head>
<BODY>
<BODY background="images/back.gif">

<%dim admin,action,checktext,pass3
admin=request("admin")
action=request("edit")
if action<>"ok" then
%>

	<table width="98%" border="1" 
		style="border-collapse: collapse; border-style: dotted; border-width: 0px"  
		bordercolor="#333333" cellspacing="0" cellpadding="2">
	<form name="thisForm" action=admin_pass.asp method=post  onsubmit="return CheckForm()">
	<tr class=backs><td colspan=2 class=td height=18>����Ա�����޸�</td></tr>
	<tr><td width=20% align=right height="18">����Ա����</td>
		<td><input type="text" readonly name="admin" size="20" value="<%=admin%>">
  &nbsp;&nbsp; <img src=../images/memo.gif alt="����Ա���Ʋ����޸�<br>������¹���Ա��ɾ���Ϲ���Ա"> </td></TR>
	<tr><td width=20% align=right height="18">�������������</td>
		<td><input type="password" name="Pass0" size="20"> </td></tr>
	<tr><td width=20% align=right height="18">��������������</td>
		<td><input type="password" name="Pass1" size="20"> 
  &nbsp;&nbsp; <img src=../images/memo.gif alt="���볤�ȣ�8-16λ<br>����ʹ�����ֺ���ĸ���"> </td></TR>
	<tr><td width=20% align=right height="18">����ȷ��������</td>
		<td><input type="password" name="Pass2" size="20">
  &nbsp;&nbsp; <img src=../images/memo.gif alt="ȷ��������"> </td></TR>
	<tr><td colspan=2><input type="submit" name="Submit" value="ȷ���޸�">
		<input type="hidden" name="edit" value="ok"></td></tr>
	</form>
	</table>
<%   
else
	

	if trim(Request("pass0"))="" or trim(Request("pass1"))="" or trim(Request("pass2"))="" then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ���д������������д������������룡');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if trim(Request("pass1"))<>trim(Request("pass2")) then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ���������������벻����������������룡');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if len(trim(Request("pass1")))<8 then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ��������������̫���ˣ�Ҫ�󳤶�Ϊ8-16λ������ʹ�����ֺ���ĸ��ϣ�');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

 Call OpenConn 
	Set rs = conn.Execute("select * from DNJY_admin where username='"&admin&"'")   
	if rs("password")<>md5(Request("pass0")) then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ������벻��ȷ��������������룡');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from DNJY_admin where username='"&admin&"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		response.write "<script language='javascript'>"
		response.write "alert('����ʧ�ܡ���δ��½����������ʱ��������������ݿ����޴˼�¼��');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
	else
		rs("password")=md5(Request("pass1"))
		rs.update

	end if
	set rs=nothing
Call CloseDB(conn)
	response.write "<script language='javascript'>"
	response.write "alert('������³ɹ��������μ�ס���������룡��\n\n���ڽ��˳��������ģ��������������µ�¼��');"
	response.write "location.href='exit.asp';"			
	response.write "</script>"

end if

%>
