<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=../Include/Function.asp-->
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
<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<script language="JavaScript" type="text/JavaScript">
function checkBig()
{
  if (document.form1.NewClassName.value=="")
  {
    alert("��Ŀ���Ʋ���Ϊ�գ�");
    document.form1.NewClassName.focus();
    return false;
  }
  if (document.form1.showid.value=="")
  {
    alert("��Ų���Ϊ�գ�");
    document.form1.showid.focus();
    return false;
  }
  
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�޸���Ŀ����</title>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
.style1 {color: #FFFFFF;
	font-weight: bold;
}
.style2 {color: #FFFFFF}
-->
</style>
</head>
<%Dim ClassID,Action,NewClassName,OldClassName,showid,rs1
ClassID=ReplaceUnsafe(Request.Form("ClassID"))
if ClassID="" then ClassID=ReplaceUnsafe(Request.QueryString("ClassID"))
Action=ReplaceUnsafe(Request("Action"))
NewClassName=ReplaceUnsafe(Request("NewClassName"))
OldClassName=ReplaceUnsafe(Request("OldClassName"))
showid=ReplaceUnsafe(Request("showid"))
if ClassID="" then
  response.Redirect("usernews_ClassManage.asp")
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "Select * From DNJY_userNewsClass where ID=" & CLng(ClassID),conn,1,3
if rs.Bof and rs.EOF then
	response.write"<SCRIPT language=JavaScript>alert('����Ŀ�����ڣ�');"
	response.write"javascript:history.go(-1)</SCRIPT>"
	response.End
else

	if Action="Modify" then
		   if NewClassName="" then
		   		response.write"<SCRIPT language=JavaScript>alert('��Ŀ������Ϊ�գ�');"
				response.write"javascript:history.go(-1)</SCRIPT>"
			    
			    response.End
			
	        end if
			
			if not IsNumeric(showid) then
		   		response.write"<SCRIPT language=JavaScript>alert('���Ӧ��Ϊ���֣�');"
				response.write"javascript:history.go(-1)</SCRIPT>"
			    
			    response.End
			
	        end if
			
			'�޸ĺ����Ŀ���������ظ�
			Set rs1=Server.CreateObject("Adodb.RecordSet")
			rs1.open "Select * From DNJY_userNewsClass Where ClassName='" & NewClassName & "' and id<>"&ClassID,conn,1,1
			if not (rs1.bof and rs1.EOF) then
				response.write"<SCRIPT language=JavaScript>alert('��Ŀ�����Ѿ����ڣ�');"
				response.write"javascript:history.go(-1)</SCRIPT>"
				
				response.End
			end if
			rs1.close
			set rs1=nothing
			
			rs("ClassName")=NewClassName
			rs("showid")=showid
			rs.update
			rs.Close
	     	set rs=Nothing
			if NewClassName<>OldClassName then				
			conn.execute "Update DNJY_userNews set ClassName='" & NewClassName & "' where ClassID=" & ClassID
     		end if		
			conn.close
			set conn=nothing
     		Response.Redirect "usernews_ClassManage.asp" 
		
	 else

%>
<body>
<div align="center"><br>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">
  <TR> 
    <TD height="5" colspan="12" bgcolor="#799AE1" align="center"><a href="usernews_ClassManage.asp"><strong><font color="#FF0000"><u>��Ŀ����</u></font></strong></a></TD>
  </TR>

      <form name="form1" method="post" action="usernews_ClassModify.asp" onSubmit="return checkBig()">
  <table border="0" cellspacing="1" align="center" class="list1">
    <tr>
      <th colspan="2" ><div align="center">�޸���Ŀ����</div></th>
    </tr>
    <tr>
      <td ><div align="right">��ĿID��</div></td>
      <td ><%=rs("ID")%>
        <input name="ClassID" type="hidden" id="ClassID" value="<%=rs("ID")%>">
        <input name="OldClassName" type="hidden" id="OldClassName" value="<%=rs("ClassName")%>"></td>
    </tr>
    <tr>
      <td ><div align="right">��Ŀ���ƣ�</div></td>
      <td ><input name="NewClassName" type="text" id="NewClassName" value="<%=rs("ClassName")%>" size="20" maxlength="30"></td>
    </tr>
    <tr>
      <td >��ʾ��ţ�</td>
      <td ><input name="showid" type="text" id="showid" value="<%=rs("showid")%>" size="6" maxlength="30">
      (����,С����ǰ��)</td>
    </tr>
    <tr>
      <td >&nbsp;</td>
      <td align="center" ><div align="left" >
          <input name="Action" type="hidden" id="Action" value="Modify">
          <input style="height:20px;" name="Save" type="submit" class="button" id="Save" value=" �� �� ">
      </div></td>
    </tr>
  </table>
  
  </form>
  <p><br>
    <br>
  </p>
</div>
</body>
<%
end if
end if
Call CloseRs(rs)
conn.close
set conn=nothing
%>
</html>
