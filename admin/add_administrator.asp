<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
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
<%call checkmanage("01")%>
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
//��֤��Ա�ʺ�����ȷ��
function checkeid(username){
var str=username;
//��JavaScript�У�������ʽֻ��ʹ��"/"��ͷ�ͽ���������ʹ��˫����
var Expression=/^[\w\u0391]+$/; //֧�ִ�Сд��ĸ�����ֺ��»����ַ�
var objExp=new RegExp(Expression);
if(objExp.test(str)==true){
return true;
}else{
return false;
}
} 
// --></script>
<script language="javascript" type="text/javascript">
<!--
function CheckForm()
{

        if(document.thisForm.username.value.length<1)
	{
	    alert("��¼�ʺŲ���Ϊ��!");
		document.thisForm.username.focus()
	    return false;
	}
    if (!checkeid(thisForm.username.value))
	{
    alert("��¼�ʺŽ�֧�ִ�Сд��ĸ�����ֺ��»����ַ�������������!");
    document.thisForm.username.focus();
    return false;
    }
	  if ((document.thisForm.password.value.Length()>16) || (document.thisForm.password.value.Length()<8)) //�ֽ������ƣ����������
     {
	 alert("���볤��Ҫ��8��16�ֽڣ����������룡");
	  document.thisForm.password.focus()
	  return false
  }
	    if(document.thisForm.password2.value.length<1)
	{
	    alert("ȷ�����벻��Ϊ��!");
		document.thisForm.password2.focus()
	    return false;
	}
	   if(document.thisForm.password2.value!=document.thisForm.password.value)
        {
            alert("�����������벻һ��!");
			document.thisForm.password2.focus()
            return false;
        }

  
}
//-->
</SCRIPT>
<div align="center">
  <form name="thisForm" method="POST" action="add_administrator_save.asp" onsubmit="return CheckForm()">
  <center>
  <table style="BORDER-COLLAPSE: collapse" borderColor="#ADAED6" height="68" cellSpacing="0" cellPadding="0" width="100%" border="1">
    <tr>
      <td bgColor="#BDBEDE" height="28">
      <p align="left" class="style1"><font size="3"><font size="2">&nbsp;</font></font>���ӹ���Ա</td>
    </tr>
    <tr>
      <td height="25" align="middle" bgcolor="#FFFFFF" style="BORDER-BOTTOM: medium none"><div align="left"><br>
          <span class="style1"><font size="3"><font size="2">&nbsp;</font></font></span><span class="style1"><font size="3"><font size="2">&nbsp;<font size="3"><font size="2">&nbsp;</font></font></font></font></span>�� �� ����
            <input name="username" size="15" maxlength="15">
      </div></td>
    </tr>
    <tr>
      <td height="25" align="middle" bgcolor="#FFFFFF" style="BORDER-TOP: medium none; BORDER-BOTTOM: medium none"><div align="left"><span class="style1"><font size="3"><font size="2">&nbsp;</font></font></span><span class="style1"><font size="3"><font size="2">&nbsp;</font></font></span><span class="style1"><font size="3"><font size="2">&nbsp;</font></font></span>��&nbsp; &nbsp;&nbsp;�룺
            <input type="password" name="password" size="15">
      </div></td>
    </tr>
    <tr>
      <td height="25" align="middle" bgcolor="#FFFFFF" style="BORDER-TOP: medium none; BORDER-BOTTOM: medium none"><div align="left"><span class="style1"><font size="3"><font size="2">&nbsp;</font></font></span><span class="style1"><font size="3"><font size="2">&nbsp;</font></font></span><span class="style1"><font size="3"><font size="2">&nbsp;</font></font></span>����ȷ�ϣ�
            <input type="password" name="password2" size="15">
      </div></td>
    </tr>
    <tr>
      <td height="25" bgcolor="#FFFFFF" style="BORDER-TOP: medium none"><p>      
        <table width="20%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><div>
              <input type="submit" value="ȷ�����" name="B1">
            </div></td>
            <td><div align="center">
              <input type="reset" value="ȡ�����" name="B1">
            </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  </center>
  </form>
</div>
<div align="center">
<%Call OpenConn
dick=request("dick")
if dick="del" then
call del()
Else
Dim username,manage
username=request("username")
if username="" then username=request.cookies("administrator")%>
</div>
<div align="center">
  <center>

<table width="100%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#335078" bordercolordark="#FFFFFF" bgcolor="#C5D5E4"  style="font-size:12px; letter-spacing:1px;">
<TR>
	<TD width="40%">  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">
    <tr bgcolor="#BDBEDE">
      <td height="28" align="center"><font color="#990000">���</font></td>
      <td height="28" align="center"><font color="#990000">�����ʺ�</font></td>      
      <td height="28" align="center"><font color="#990000">�޸�����</font></td>
	  <td height="28" align="center"><font color="#990000">����Ȩ��</font></td>
      <td height="28" align="center"><font color="#990000">��������</font></td>
	  <td height="28" align="center"><font color="#990000">ɾ��</font></td>	  
    </tr>
<%dim i,id,dick,deladm
i=1
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select id,username,manage from [DNJY_admin]" 
rs.open sql,conn,1,1
do while not rs.eof
manage=rs("manage")%>
    <tr>
      <td width="30" height="24" align="center" bgcolor="#FAFAFA"><%=i%></td>
      <td width="100" height="24" align="center" bgcolor="#FFFFFF"><div align="left"><%if username=rs("username") then
		response.write "<a href=add_administrator.asp?username="&rs("username")&"><font color=red><b>"&rs("username")&"</b></font></a>"
		else
		response.write "<a href=add_administrator.asp?username="&rs("username")&">"&rs("username")&"</a>"
	    end if%></div></td>      
	  <td width="50" height="24" align="center" bgcolor="#FAFAFA"><a href="admin_pass.asp?admin=<%=rs("username")%>">�޸�</a></div></td>
      <td width="60" height="24" align="center" bgcolor="#FFFFFF"><div align="left"><%
		response.write "<a href=add_administrator.asp?username="&rs("username")&">����Ȩ��</a>"%></div></td> 
	<td width="60" height="24" align="center" bgcolor="#FAFAFA"><a href="#" onClick="window.open('user_editdata.asp?username=<%=rs("username")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=550,height=650,left=300,top=100')">�޸�����</a></td>
    <td width="30" height="24" align="center" bgcolor="#FAFAFA"><img border=0 src=images/delete.gif alt="ɾ���˹���Ա" style="cursor:hand" onclick="{if(confirm('�ò������ɻָ���\n\nȷʵҪɾ���������Ա�� ')){location.href='add_administrator.asp?dick=del&deladm=<%=rs("username")%>';}}"></td>

    </tr>
<%
i=i+1
rs.movenext
loop
Call CloseRs(rs)%>
</table></TD>
	<TD width="60%"><table border=0 width=100%>
	<form action="add_administrator_save.asp" method=post name=manage>
	<tr><td>����Ա<font color=red><b><%=username%></b></font> �Ĺ���Ȩ�ޣ�</td></tr>
	<%
	   Set rs = conn.Execute("select * from DNJY_admin where username='"&username&"'")
		  if rs.eof and rs.bof then
			response.write "<script language='javascript'>"
			response.write "alert('û�д˹���Ա��');"
			response.write "location.href='javascript:history.go(-1)';"
			response.write "</script>"
			response.end
		  end if
 		  manage=rs("manage")
		%>
		<tr><td><%if username=request.cookies("administrator") then%><input type=checkbox value='yes' name=manage01 DISABLED><%Else%><input type=checkbox value='yes' name=manage01 
			<%if instr(manage,"01")>0 then%>checked<%end if%>><%end if%><font color=FFFF00>����Ȩ��</font>&nbsp&nbsp;&nbsp;&nbsp;<input type=checkbox value='yes' name=manage02 
			<%if instr(manage,"02")>0 then%>checked<%end if%>>ϵͳ����&nbsp&nbsp;&nbsp;&nbsp;<input type=checkbox value='yes' name=manage03 
			<%if instr(manage,"03")>0 then%>checked<%end if%>>��Ա����&nbsp&nbsp;&nbsp;&nbsp;<input type=checkbox value='yes' name=manage04 
			<%if instr(manage,"04")>0 then%>checked<%end if%>>��Ա����&nbsp&nbsp;&nbsp;&nbsp;<input type=checkbox value='yes' name=manage05 
			<%if instr(manage,"05")>0 then%>checked<%end if%>>��Ϣ����<br>
			<tr><td><input type=checkbox value='yes' name=manage06 
			<%if instr(manage,"06")>0 then%>checked<%end if%>>���Ź���&nbsp&nbsp;&nbsp;&nbsp;<input type=checkbox value='yes' name=manage07 
			<%if instr(manage,"07")>0 then%>checked<%end if%>>���Թ���&nbsp&nbsp;&nbsp;&nbsp;<input type=checkbox value='yes' name=manage08 
			<%if instr(manage,"08")>0 then%>checked<%end if%>>�ʼ���־&nbsp&nbsp;&nbsp;&nbsp;<input type=checkbox value='yes' name=manage09 
			<%if instr(manage,"09")>0 then%>checked<%end if%>>վ�ڶ���&nbsp&nbsp;&nbsp;&nbsp;<input type=checkbox value='yes' name=manage10 
			<%if instr(manage,"10")>0 then%>checked<%end if%>>�ɼ�����<br>
			<tr><td><input type=checkbox value='yes' name=manage11 
			<%if instr(manage,"11")>0 then%>checked<%end if%>>ģ�����&nbsp&nbsp;&nbsp;&nbsp;<input type=checkbox value='yes' name=manage12 
			<%if instr(manage,"12")>0 then%>checked<%end if%>>��������&nbsp&nbsp;&nbsp;&nbsp;<input type=checkbox value='yes' name=manage13 
			<%if instr(manage,"13")>0 then%>checked<%end if%>>������&nbsp&nbsp;&nbsp;&nbsp;<input type=checkbox value='yes' name=manage14 
			<%if instr(manage,"14")>0 then%>checked<%end if%>>��ȫ����&nbsp&nbsp;&nbsp;&nbsp;<input type=checkbox value='yes' name=manage15 
			<%if instr(manage,"15")>0 then%>checked<%end if%>>����ͼƬ</td></tr>	

		</table></TD>
			<tr><td colspan=2>
		<input type="submit" name="Submit" value="��������">
		<input type="hidden" name="edit" value="ok">
		<input type="hidden" name="username" value="<%=username%>">
		<%if username=request.cookies("administrator") then%><input type="hidden" name="manage01" value="yes"><%end if%>
		</form>
	</td></tr>
</TR>
</TABLE>

<%sub del()
id=request("id")
deladm=request("deladm")

if deladm=request.cookies("administrator") then
			response.write "<script language='javascript'>"
			response.write "alert('������ɾ���Լ�-_-��\n\n��ɾ����������һ������Ա�ٽ�����');"
			response.write "location.href='add_administrator.asp';"
			response.write "</script>"
			response.end
		else
set rs = Server.CreateObject("ADODB.RecordSet")
sql="delete from [DNJY_admin] where username='"&deladm&"'"
rs.open sql,conn,1,3

response.write "<script language=JavaScript>" & chr(13) & "alert('ɾ���ʺųɹ���');" & "window.location='add_administrator.asp'" & "</script>"
response.end
end if
end sub
set rs=nothing
end if
Call CloseDB(conn)
%>
  </table>
  </center>
</div>
