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
<script language=javascript src=Include/mouse_on_title.js></script>
<style type="text/css">
<!--
.style1 {color: #FF0000}
.style5 {color: #FF0000; font-weight: bold; }
-->
</style>
<script language="javascript">
<!--
function chkinput(){
if (document.msgadd.name.value==0)
	{
	  alert("������ջ�ԱID�ţ�")
	  document.msgadd.name.focus()
	  return false
	 }
if (document.msgadd.neirong.value==0)
	{
	  alert("������Ϣ���ݣ�")
	  document.msgadd.neirong.focus()
	  return false
	 }
if (document.msgadd.validate_codee.value==0)
	{
	  alert("������֤�룡")
	  document.msgadd.validate_codee.focus()
	  return false
	 }
}
// --></script>
</head>

<body background="img/bg1.gif" leftmargin="0" topmargin="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="980" height="426">
    <tr>
      <td width="214" height="354" valign="top"bgcolor="#FFFFFF"><div align="center">
      
 <!--#include file=userleft.asp--></div></td>
 <td width="5">&nbsp;</td>
 <td width="760" height="354" valign="top" bgcolor="#FFFFFF" class="ty1">
 <%Dim rsmess,pages,allPages,page,neirong,riqi,isnew,name,fname,maxlength,rsuser,username2,huifu
if request.cookies("dick")("username")="" then
response.Write "<center>���ȵ�¼</center>"
response.End
end if
username=request.cookies("dick")("username")
fname=request.Form("fname")
username2=request.Form("username2")
name=request("name")
id=request.Form("id")
%>
<%If request("action")="msgsave" Then
    if CInt(trim(request.form("validate_codee")))<>CInt(Session("ttpSN")) then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ���֤�벻�ԣ�');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
    Session("ttpSN")=""
    response.end
    end if
    if request.form("name")="" then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ��û�ID��û�����ܸ���������?');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end If
    if request.form("name")="ALL"  Or request.form("name")="all" then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ�������ϵͳר���û���ALL��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end If
	set rsuser=server.createobject("adodb.recordset")
	rsuser.Open "select username from [DNJY_user] where username='"&request("name")&"'",conn,1,1	
	if rsuser.eof and rsuser.bof then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ�û���ҵ�"&request("name")&"��Ա������������ύ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if	
    if request.form("neirong")=""  then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ�û����дҪ���͵����ݣ�');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	rsuser.close
    set rsuser=nothing
	response.end
	end if

	if len(request.form("neirong"))>300  then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ�����д������̫���ˣ�����������ύ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	sql="select * from DNJY_Message"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
rs.AddNew
  	rs("name")=name
	neirong=request.form("neirong")
	neirong=replace(neirong,"=======","<font color=darkgray>======")
  	rs("neirong")=HTMLEncode(neirong)
  	rs("riqi")=now()
 	rs("fname")=request.cookies("dick")("username")
rs.Update
Call CloseRs(rs)

	response.write "<script language='javascript'>"
	response.write "alert('�����ɹ�������"&request("name")&"����һ��վ����Ϣ��');"
	response.write "location.href='messs.asp?"&C_ID&"';"
	response.write "</script>"
	response.end
conn.close
set conn=Nothing
else%>
<form name=msgadd method=post action=?action=msgsave&<%=C_ID%> onsubmit="return chkinput()">
<%if id<>"" then
Set rs = conn.Execute("select * from DNJY_Message where id="&id) 
if not (rs.eof and rs.bof) then
'����ͨ���ظ���Ϣ�ķ�ʽ͵�����˵���Ϣ
if rs("name")=request.Cookies(ttSavename&"_shop")("username") then
huifu=vbcrlf&vbcrlf&vbcrlf&vbcrlf&"======="&rs("fname")&"��"&rs("riqi")&"���͵���Ϣԭ��======"&vbcrlf&replace(rs("neirong"),"<font color=darkgray>","")
end if
end if
set rs=nothing
end if
%>
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
<table width="760" border="0" cellpadding="4" cellspacing="0" bordercolor="#C0C0C0" align=center bordercolordark="#FFFFFF">
<tr><td>����վ����Ϣ��</tr>
	<tr><td width=20%  height="18">�� �� ˭&nbsp;
	<input type=text 
	<%
	if request("fname")<>"" then response.write "value='"&fname&"'"
	if request("username2")<>"" then response.write "value='"&username2&"'"
	if request("name")<>"" then response.write "value='"&name&"'"
	%>"
	name="name" size=18 maxlength=16> ������д��վע���Ա��ID��</td></tr>
<tr><td>
  <textarea rows="15" name="neirong" cols="75" style="BORDER: darkgray 1px solid; FONT-SIZE: 8pt; FONT-FAMILY: verdana ; overflow:auto;" onkeyUp="textLimitCheck(this,300);"><%=huifu%></textarea><br>��<%=300%>���ַ� ������ <font color="#CC0000"><span id="messageCount">0</span></font> ���� 
&nbsp;&nbsp;&nbsp;����htm����</td></tr>
<tr><td>��ʽ��֤��<img src="../tt_getcodee.asp" alt="��֤��,�������?����ˢ����֤��" height="10" style="cursor : pointer;" onClick="this.src='../tt_getcodee.asp?t='+(new Date().getTime());" /><input type="text" name="validate_codee" size="4" onkeyup="value=value.replace(/[^\d|^\.]/g,'')"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))">&nbsp;&nbsp; <img src=../images/memo.gif alt="��֤��,�������?����ˢ����֤��"></td></tr>
<tr><td height=50 align=center><input type="submit" value="��������" name="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" value="������д"><input name="send" type="hidden" value="ok"></td></tr>
</form></td>
</table>      
    </tr>
    <tr>
      <td height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
  <%End if%>
</div>
</body>
</html>