<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/tt_auto_lock.asp" -->
<%response.buffer="True"%>
<!--#include file="../Include/md5.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<script language="JavaScript">

<!--
function get_onsubmit() {
if (document.get.userid.value=="")
	{
	  alert("���趨���ĵ�½����")
	  document.get.userid.focus()
	  return false
	 }	 	 
}
// -->
</script>
<meta http-equiv="Content-Language" content="zh-cn">
<link href="/<%=strInstallDir%>css/css.css" rel="stylesheet" type="text/css">
<title>�����һ�</title>
<style>
<!--
.style1 {
	font-size: 16px;
	font-family: "����";
}
-->
</style>

</head>
<% dim fpass1,fpass,strSql,username
Call OpenConn
if request("action")="edit" then
username=trim(request.form("username"))
fpass1=trim(request.form("fpass1"))
fpass=md5(fpass1)
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from [DNJY_user] where  username='"&username&"'",conn,1,3
if rs.eof or rs.bof then
response.write "<li>��������"
Call CloseRs(rs)
Call CloseDB(conn)
response.end
end if
rs("password")=fpass
rs.update
Call CloseRs(rs)
Call CloseDB(conn)
response.write "<script>alert('�����޸ĳɹ������ס��');history.back();</Script>"
response.end
end if%>

<%dim answer1,answer
if Request.form("submit")<>"" then 
username=request.form("username") 'ҳ��֮�䴫�ݲ���
answer1= trim(Request.form("answer1"))
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_user] where username='"&UserName&"'"
rs.open sql,conn,1,1
answer =rs("answer")
answer1=md5(answer1)
Call CloseRs(rs)
if answer<>answer1 Or answer="d41d8cd98f00b204e9800998ecf8427e" then
session("gpw_error")=session("gpw_error")+1
response.write "<script>alert('��Ĵ𰸴����뷵�ؼ�飡���Ѵ�������"&session("gpw_error")&"�Σ����޽�����IP�޷����ʣ�');history.back();</Script>"
Call CloseDB(conn)
response.end
else
end if
end if

Call CloseDB(conn)%>
<center>
<script language="VBScript">
<!--
function checkaddurl()
fpass1=urlform.fpass1.value
if urlform.fpass1.value="" or len(fpass1)<5 or len(fpass1)>12 then 
msgbox"�����������루��5--12���ַ���^_^"
urlform.fpass1.focus
elseif urlform.fpass1.value<>"" and urlform.fpass2.value<>urlform.fpass1.value then 
msgbox"ȷ�����벻��ͬ��^_^"
urlform.fpass2.focus
else
urlform.submit
end if

end function
-->
</script>
<body topmargin="0" leftmargin="0">
<%If request.form("username")<>"" then%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="400" id="AutoNumber1" height="147">
    <tr> 
      <td width="500" height="25" bgcolor="#FF9900"> <p align="center">&nbsp; ..:::�����������������������:::...</td>
    </tr>
  <tr> 
    <td width="500" height="122"> 
      <table border="0" width="100%" id="table1" height="100">
  <form name="urlform" method="post" action="RetakePassword_c.asp?action=edit" onsubmit="return(search4())" >
    
  
    <tr valign="baseline"> 
      <td nowrap align="right"  class=tdc><span class="style1">���������룺</span></td>
      <td  class=tdc> 
        <input type="password" name="fpass1" value="" size="20" maxlength="12">��5-12�ַ��ڣ�
      </td>
    </tr>
     <tr valign="baseline"> 
      <td nowrap align="right"  class=tdc><span class="style1">����ȷ�ϣ�</span></td>
      <td  class=tdc> 
        <input type="password" name="fpass2" value="" size="20" maxlength="12">��5-12�ַ��ڣ� 
      </td>
    </tr>
     
        <tr valign="baseline"> 
      <td nowrap align="right"  class=tdc>��</td>
      <td  class=tdc> 
	   <input type="hidden" name="username" value="<%=trim(request.form("username"))%>">
        <input type="button" name="sss" value="�ύ" onclick=checkaddurl()>
              �� 
              <input type="button" name="aaa" value="ȡ��" onClick="javascript:window.close()"> 
      </td>
		</tr></form>
		</table>
	</td>
  </tr>
  </table>
  <%else%>
<table border="0" width="400" id="table2" bgcolor="#FF9900">
	<tr>
		<td>��������</td>
	</tr>
</table>
<%End if%>
<table border="0" width="400" id="table2" bgcolor="#FF9900">
	<tr>
		<td>��</td>
	</tr>
</table>