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
<%call checkmanage("03")%>
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
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
<script>
function loadThreadFollow(t_id,b_id){
	var targetImg =eval("document.all.followImg" + t_id);
	var targetDiv =eval("document.all.follow" + t_id);
	
	if ("object"==typeof(targetImg)){
		if (targetDiv.style.display!='block'){
			targetDiv.style.display="block";
			
		}else{
			targetDiv.style.display="none";
		}
	}
}
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall' )
       e.checked = form.chkall.checked;
    }
  }
//-->
</script>
<script language=javascript>
//�����͹���ظ���ʼ
<!--
function DoEmpty(params)
{
if (confirm("���Ҫɾ��������¼��ɾ�����޷��ָ���"))
window.location = params ;
}
function test()
{
  if(!confirm('�Ƿ�ȷ�����������������������ָܻ���')) return false;
}
-->
</script>
<%Call OpenConn
Dim id,str2,i
id=request("selectedid")
dick=request("dick")
If dick="yz" Then
conn.execute "UPDATE DNJY_Financial SET admincl=1 WHERE id="&id&"" 
End if
If dick="delyz" Then
conn.execute "UPDATE DNJY_Financial SET admincl=0 WHERE id="&id&"" 
End If

If dick="del" Then
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" &"window.location='admin_Amount.asp'" & "</script>"
response.end
end if
str2=split(id,",")
set rs=server.createobject("adodb.recordset")
for i=0 to ubound(str2)
sql="delete from [DNJY_Financial] where id="&cstr(str2(i))
rs.open sql,conn,1,3
next
response.write "<script language=JavaScript>" & chr(13) & "alert('ɾ���ɹ���');" &"window.location='admin_Amount.asp'" & "</script>"
response.end
Call CloseRs(rs)
End If%>
<FORM name=thisForm action="admin_Amount.asp?dick=del" method=POST>
<div align="center">
  <center>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%" height="130">
<TR>
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">�� �� �� �� �� ��</FONT></TD>
 </TR>
  <tr>
    <td height="28" colspan="12" align="center"><div align="left"><span class="style1"><a href="admin_Amount.asp"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font></font></span></a></span><span class="style2"><a href="admin_Amount.asp">����ȫ������</a>&nbsp;&nbsp; <a href="admin_Amount.asp?dick=2">δ��ʾ����</a>&nbsp;&nbsp; <a href="admin_Amount.asp?dick=1">����ʾ����</a>&nbsp;&nbsp; <a href="#" onClick="window.open('admin_Amount_add.asp','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=480,height=400,left=300,top=100')"><b>���������<b></a>&nbsp;&nbsp;<input type=button value=ˢ�� onclick="window.open('admin_amount.asp','_self')"> </span></div></td>
    </tr>
  <tr bgcolor="#eeeeee">
   <td width="5%" height="24" align="center" bgcolor="#eeeeee">ѡ��</td>
    <td width="5%" height="24" align="center" bgcolor="#eeeeee">���</td>
    <td width="8%" height="24" align="center">ע���û�</td>    
    <td width="8%" height="24" align="center">�û�����</td>
    <td width="5%" height="24" align="center">ҵ����</td>
    <td width="5%" height="24" align="center">ҵ������</td> 
    <td width="20%" height="24" align="center">����ʱ��</td>                
    <td width="20%" height="24" align="center">��ע</td>
    <td width="5%" height="24" align="center">����</td> 	
    <td width="5%" height="24" align="center">������</td> 	
    <td width="5%" height="24" align="center">ǰ̨��ʾ</td>
    <td width="5%" height="24" align="center">ɾ��</td>
    </tr>
<%
dim k,dick
k=0
dim ThisPage,Pagesize,Allrecord,Allpage
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end if
dick=request("dick")
set rs=server.createobject("adodb.recordset")
Select Case dick
Case "1"
sql = "select * from DNJY_Financial where admincl=1 order by id "&DN_OrderType&""
Case "2"
sql = "select * from DNJY_Financial where admincl=0 order by id "&DN_OrderType&""
Case Else
sql = "select * from DNJY_Financial order by id "&DN_OrderType&""
End Select
rs.open sql,conn,1,1
if rs.eof then
response.write "<table>��û�м�¼��</table>"
response.end
end if
dim uid
rs.Pagesize=20
Pagesize=rs.Pagesize
Allrecord=rs.Recordcount
Allpage=rs.Pagecount
if ThisPage<1 then                           
ThisPage=1
end if
On Error Resume Next
rs.move (ThisPage-1)*Pagesize
k=0
do while not rs.eof
uid=rs("id")
%>
  <tr bgcolor="#FFFFFF">
   <td width="5%" height="25" align="center" bgcolor="#FFFFFF">
   <input type="checkbox" name="selectedid" value="<%=trim(rs("id"))%>"></td>
    <td  height="24" align="center" bordercolor="#999999" bgcolor="#FAFAFA" style="border-bottom-style: solid; border-bottom-width: 1"><div align="center"><%=k+1%></div></td>     
    <td  height="20" align="center" bordercolor="#999999" style="border-bottom-style: solid; border-bottom-width: 1"><a href=users_List.asp?T1=<%=rs("username")%>&dick=1><font color="#FF0000"><%=rs("username")%></font></a></td> 
    <td height="20" align="center" bordercolor="#999999" style="border-bottom-style: solid; border-bottom-width: 1"><font color="#FF0000"><%=rs("aliname")%></font></td>
    <td  height="24" align="center" bordercolor="#999999" style="border-bottom-style: solid; border-bottom-width: 1"><font color="#FF0000"><%=rs("ywje")%></font></td>
    <td  height="24" align="center" bordercolor="#999999" bgcolor="#FAFAFA" style="border-bottom-style: solid; border-bottom-width: 1"><%If rs("ywlx")="֧��" then%><font color="#FF0000"><%elseif rs("ywlx")="����" then%><font color="#0000FF"><%End if%><%=rs("ywlx")%></font></td>
    <td  height="24" align="center" bordercolor="#999999" bgcolor="#FAFAFA" style="border-bottom-style: solid; border-bottom-width: 1"><%=rs("addTime")%></td>	
   <td  height="24" align="center" bordercolor="#999999" bgcolor="#FAFAFA" style="border-bottom-style: solid; border-bottom-width: 1"><%=rs("czbz")%></td>
   <td  height="24" align="center" bordercolor="#999999" style="border-bottom-style: solid; border-bottom-width: 1"><a href="#" onClick="window.open('admin_Amount_add.asp?dick=uyz&username=<%=rs("username")%>&aliname=<%=rs("aliname")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=480,height=410,left=300,top=100')"><font color="#FF0000">����</font></a></td> 
   <td  height="24" align="center" bordercolor="#999999" bgcolor="#FAFAFA" style="border-bottom-style: solid; border-bottom-width: 1"><%=rs("czz")%></td> 
    <td  height="24" align="center" bordercolor="#999999" style="border-bottom-style: solid; border-bottom-width: 1"><%if rs("admincl")=1 then%><a href="?dick=delyz&selectedid=<%=rs("id")%>"><font color="#008000">��</font></a><%else%><a href="?dick=yz&selectedid=<%=rs("id")%>"><font color="#FF0000">��</font></a><%end if%></td>
    <td  height="24" align="center" bordercolor="#999999" style="border-bottom-style: solid; border-bottom-width: 1">
    <a href="javascript:DoEmpty('?dick=del&selectedid=<%=rs("id")%>')">ɾ��</a></td>
    </tr>
<%rs.movenext
k=k+1
if k>=Pagesize then exit do
loop
Call CloseRs(rs)
Call CloseDB(conn)%>
</table>
 <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
<tr> 
<td width="595" height="25" align="center">
  <input onClick=CheckAll(this.form) type=checkbox value=on name=chkall>
  ȫѡ��¼
  <input name='submit1' type='submit' value=' ����ɾ�� ' onClick="return test();"  style="color: #FF0000">
  ����&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;����¼ �� <font color="#CC5200"><%=Allpage%></font> ҳ �����ǵ� <font color="#CC5200"><%=ThisPage%></font> ҳ
<%
if ThisPage<2 then     
response.write "<font color=""#808080"">��ҳ</font>&nbsp;"
response.write "<font color=""#808080"">��һҳ</font>&nbsp;"     
else     
response.write "<a href=?page=1&t1="&trim(request("t1"))&">��ҳ</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&t1="&trim(request("t1"))&">��һҳ</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">��һҳ</font>&nbsp;"
response.write "<font color=""#808080"">βҳ</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&t1="&trim(request("t1"))&">��һҳ</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&t1="&trim(request("t1"))&">βҳ</a>&nbsp;"     
end if
%></td>
</tr>
          </table>
  </center>
</div>
</form>
