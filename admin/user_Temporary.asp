<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
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
<%Call OpenConn
Dim id,str2,i
id=request("selectedid")
dick=request("dick")

If dick="del" Then
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" &"window.location='user_Temporary.asp'" & "</script>"
response.end
end if
str2=split(id,",")
set rs=server.createobject("adodb.recordset")
for i=0 to ubound(str2)
sql="delete from [DNJY_usertemp] where id="&cstr(str2(i))
rs.open sql,conn,1,3
next
response.write "<script language=JavaScript>" & chr(13) & "alert('ɾ���ɹ���');" &"window.location='user_Temporary.asp'" & "</script>"
response.end
Call CloseRs(rs)
End If%>
<FORM name=thisForm action="user_Temporary.asp?dick=del" method=POST>
<div align="center">
  <center>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%" height="130">
<TR>
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">��ʱ��Ա����</FONT></TD>
 </TR>
 <TR>
    <TD height="20" bgcolor="#ffffff" colspan="13">&nbsp;&nbsp;&nbsp;&nbsp;��ϵͳ�����˻�Աע���ʼ���֤�󣬻�Աע����ͨ���ʼ���֤ǰ��ע�����Ͻ������ڴˣ���ͨ���ʼ���֤�򳬹��޶�ʱ��δ�����ʼ���֤�ģ����Զ�ɾ������վ����Ա���ڴ˹۲��Աע�������Ҳ����ɾ��������</TD>
 </TR>
  <tr bgcolor="#eeeeee">
   <td width="5%" height="24" align="center" bgcolor="#eeeeee">ѡ��</td>
    <td width="5%" height="24" align="center" bgcolor="#eeeeee">���</td>
    <td width="10%" height="24" align="center">��¼�ʺ�</td>    
    <td width="10%" height="24" align="center">��Ա����</td>
    <td width="15%" height="24" align="center">��������</td> 
    <td width="10%" height="24" align="center">ע��ʱ��</td>
	<td width="10%" height="24" align="center">��Чʱ��(��:ʱ:��:��)</td>
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
sql = "select * from DNJY_usertemp order by id "&DN_OrderType&""
rs.open sql,conn,1,1
if rs.eof then
response.write "<table>��û�м�¼��</table>"
Call CloseRs(rs)
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
    <td  height="20" align="center" bordercolor="#999999" style="border-bottom-style: solid; border-bottom-width: 1"><%=rs("username")%></td> 
    <td height="20" align="center" bordercolor="#999999" style="border-bottom-style: solid; border-bottom-width: 1"><%=rs("name")%></td>
    <td  height="24" align="center" bordercolor="#999999" style="border-bottom-style: solid; border-bottom-width: 1"><%=rs("email")%></td>
    <td  height="24" align="center" bordercolor="#999999" bgcolor="#FAFAFA" style="border-bottom-style: solid; border-bottom-width: 1"><%=rs("zcdata")%></td>
    <td  height="24" align="center" bordercolor="#999999" bgcolor="#FAFAFA" style="border-bottom-style: solid; border-bottom-width: 1"><%=tmraumen_Timer((rs("zcdata")+regyxq))%></td>	
    <td  height="24" align="center" bordercolor="#999999" style="border-bottom-style: solid; border-bottom-width: 1">
    <a href="?dick=del&selectedid=<%=rs("id")%>">ɾ��</a></td>
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
  <input  type="submit" value="ɾ����¼" name="B1">
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
<%If mailqr=1 then
Dim rsdq
set rsdq=server.createobject("adodb.recordset")
sql="delete from [DNJY_usertemp] where DateDiff("&DN_DatePart_D&",zcdata,"&DN_Now&")>"&regyxq&""'ɾ��3��δ�ʼ���֤����ʱע����Ϣ
rsdq.open sql,conn,1,3
End If
%>