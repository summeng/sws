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
<%call checkmanage("05")%>
<link rel="stylesheet" type="text/css" href="../css/style.css">
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
	margin-top: 16px;
}
-->
</style>
<body>
<p align="center">
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
</script>
<SCRIPT language=JavaScript>
function showoperatealert(id)
{
		if (id==1)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="info_replyyz.asp";
			thisForm.submit();
		}
		}
		if (id==2)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="info_replydel.asp";
			thisForm.submit();
		}
		}
		if (id==3)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="info_replyyz.asp?dick=delyz";
			thisForm.submit();
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
function DoEmpty(params)
{
if (confirm("���Ҫɾ��������¼��ɾ�����޷��ָ���"))
window.location = params ;
}
function test()
{
  if(!confirm('�Ƿ�ȷ�����������������������ָܻ���')) return false;
}
//-->
    </SCRIPT>
</p>
<div align="center">
<table border="0" cellpadding="0" cellspacing="0" style="border:1px solid #CCCCCC; " width="100%" height="1">
<tr>
<td width="100%" height="24" colspan="6">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="59">
<TR>
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">�� �� �� Ϣ �� ��</FONT></TD>
 </TR>
<tr>
<td height="31">
<FORM name=sou1 action="?dick=1" method=POST style="line-height: 150%; margin-top: 0; margin-bottom: 0">
<p align="center">���û��ʺţ�<input type="text" name="T1" size="10">
<input type="submit" value="����" name="B1">
</form>
</td>
<td>
<FORM name=sou3 action="?dick=3" method=POST style="line-height: 150%; margin-top: 0; margin-bottom: 0">
<p align="center">���ؼ��ʣ�<input type="text" name="T3" size="10">
<input type="submit" value="����" name="B1">
</form>
</td>
</tr>
<tr>
<td width="864" colspan="2" height="28">
<p align="center"><a href="?dick=4">δ(<font color="#FF0000">��</font>)��֤�ظ�</a>---<a href="?dick=5">��(��)��֤�Ļظ�</a>---<a href="?dick="><font color="#FF0000">��ʾȫ���ظ�</font></a></td>
</tr>
</table>
</td>
</tr>
<FORM name=thisForm action="info_replydel.asp" method=POST>
<%Call OpenConn
dim k,dick
dim ThisPage,Pagesize,Allrecord,Allpage
k=1
set rs = Server.CreateObject("ADODB.RecordSet")
dick=trim(request("dick"))
Select Case dick
Case "1"
sql="select * from [DNJY_reply] where username='"&trim(request("T1"))&"' order by hfsj "&DN_OrderType&"" 
Case "3"
sql="select * from [DNJY_reply] where neirong like '%"&trim(request("T3"))&"%' or neirong like '%"&trim(request("T3"))&"%' order by hfsj "&DN_OrderType&"" 
Case "4"
sql="select * from [DNJY_reply] where hfyz=0 order by hfsj "&DN_OrderType&"" 
Case "5"
sql="select * from [DNJY_reply] where hfyz=1 order by hfsj "&DN_OrderType&""
Case Else
sql="select * from [DNJY_reply] order by hfsj "&DN_OrderType&"" 
End Select
rs.open sql,conn,1,1
if rs.eof then
response.write "<tr><td><li>û�м�¼"
response.end
end if
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end if
rs.Pagesize=30
Pagesize=rs.Pagesize
Allrecord=rs.Recordcount
Allpage=rs.Pagecount
if ThisPage<1 then                           
ThisPage=1
end if
On Error Resume Next
rs.move (ThisPage-1)*Pagesize
%>
<tr>
<td width="5%" height="25" align="center" bgcolor="#FEEEBC" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#999999">ѡ��</td>
<td width="20%" bgcolor="#FEEEBC" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#999999">�ظ�����</td>
<td width="20%" height="25" align="center" bgcolor="#FEEEBC" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#999999">�����Ϣ</td>
<td width="10%" bgcolor="#FEEEBC" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#999999">
<p align="center">�ظ���</td>
<td width="20%"  bgcolor="#FEEEBC" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#999999">
<p align="center">����</td>
<td width="10%" bgcolor="#FEEEBC" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#999999">
<p align="center">����IP</td>
<td width="5%" bgcolor="#FEEEBC" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#999999">
<p align="center">��֤</td>
<td width="5%" bgcolor="#FEEEBC" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#999999" align="center">
����</td>
</tr>
<%
dim uid
do while not rs.eof
uid=rs("xxid")
%>
<tr onmouseover="this.style.backgroundColor='#EFE8D6';return true;" onmouseout="this.style.backgroundColor='#FFFFFF';">
<td height="24" align="center" style="border-bottom:1px solid #CCCCCC; border-top-width: 1px; border-right-width:1px">
<INPUT type=checkbox value="<%=rs("id")%>" name=selectedid></td>
<td style="border-left:1px solid #CCCCCC; border-bottom:1px solid #CCCCCC; border-top-width: 1px; border-right-width:1px"><% if len(rs("neirong"))>15 then
response.write "<a href='chkcomment.asp?edit=replay&id="&rs("id")&"'>"&left(trim(rs("neirong")),13)&"...</a>"
else
response.write "<a href='chkcomment.asp?edit=replay&id="&rs("id")&"'>"&trim(rs("neirong"))&"</a>"
end if%></td>
<td height="24" style="border-bottom:1px solid #CCCCCC; border-top-width: 1px; border-right-width:1px">
<a target='_blank' href=../show.asp?id=<%=rs("xxid")%>><%=conn.Execute("Select biaoti From DNJY_info where id="&rs("xxid")&"")(0)%></a></td>
<td align="center" style="border-left:1px solid #CCCCCC; border-bottom:1px solid #CCCCCC; border-top-width: 1px; border-right-width:1px"><%if rs("username")<>"" then%><a target="_blank" href=../preview.asp?username=<%=rs("username")%>><%=rs("username")%><%else%>����<%end if%></td>
<td align="center" style="border-left:1px solid #CCCCCC; border-bottom:1px solid #CCCCCC; border-top-width: 1px; border-right-width:1px"><%=rs("hfsj")%></td>
<td align="center" style="border-left:1px solid #CCCCCC; border-bottom:1px solid #CCCCCC; border-top-width: 1px; border-right-width:1px"><%=rs("ip")%></td>
<td  align="center" style="border-left:1px solid #CCCCCC; border-bottom:1px solid #CCCCCC; border-top-width: 1px; border-right-width:1px"><%if rs("hfyz")=1 then%><a href="info_replyyz.asp?dick=delyz&selectedid=<%=rs("id")%>">��</a><%else%><b><a href="info_replyyz.asp?selectedid=<%=rs("id")%>"><font color="#FF0000">��</font></a></b><%end if%></td>
<td style="border-left:1px solid #CCCCCC; border-bottom:1px solid #CCCCCC; border-top-width: 1px; border-right-width:1px">
<p align="center"><font color="#FF0000">
<span id="followImg<%=k%>" style="CURSOR: hand" onclick="loadThreadFollow(<%=k%>,5)">
&nbsp;</span></font><a href="javascript:DoEmpty('info_replydel.asp?selectedid=<%=rs("id")%>&xid=<%=uid%>')">ɾ��</a></td>
</tr>
<%
k=k+1
rs.movenext
if k>=Pagesize then exit do
loop
Call CloseRs(rs)
Call CloseDB(conn)
%>
<tr>
<td width="100%" height="1" colspan="6" style="border-top-style: solid; border-top-width: 1">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="1">
<tr> 
<td height="25" width="500" colspan="5">
<p align="left">
��<INPUT onclick=CheckAll(this.form) type=checkbox value=on name=chkall>ѡ������<input onclick=javascript:showoperatealert(1) type="submit" value="ͨ����֤" name="B1"> 
<input onclick=javascript:showoperatealert(3) type="submit" value="ȡ����֤" name="B3">
<input name='submit1' type='submit' value=' ����ɾ�� ' onClick="return test();"  style="color: #FF0000">
</td> 

</tr>
<tr> 
<td height="25" width="800" align="right">
����&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;����¼
�� <font color="#CC5200"><%=Allpage%></font> ҳ
�����ǵ� 
<font color="#CC5200"><%=ThisPage%></font> ҳ
<%
if ThisPage<2 then     
response.write "<font color=""#808080"">��ҳ</font>&nbsp;"
response.write "<font color=""#808080"">��һҳ</font>&nbsp;"     
else     
response.write "<a href=?page=1&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&dick&">��ҳ</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&dick&">��һҳ</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">��һҳ</font>&nbsp;"
response.write "<font color=""#808080"">βҳ</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&dick&">��һҳ</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&dick&">βҳ</a>&nbsp;"     
end if
%></td>
</tr>
<tr> 
<td height="1" width="595" colspan="4">
<p align="right"></td>
</tr>
</table>
</td>
</tr>
</form>
<tr>
<td width="100%" height="1" colspan="6" style="border-top-style: solid; border-top-width: 1"></td>
</tr>
</table>
</div>