<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"--> 
<!--#include file="usercookies.asp"-->
<!--#include file=top.asp-->

<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%dim sc,ii,e,str2,xxid,rs0,sql0,rsk,sqlk
sc=Request("sc")%>
</head>
<SCRIPT language=JavaScript>
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall' )
       e.checked = form.chkall.checked;
    }
  }
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
//-->
</SCRIPT>
<body topmargin="0">

<div align="center">
<table width="980" id="table1" bgcolor="#FFFFFF"  cellspacing="1">

<tr>
<td width="158" height="144" valign="top"><!--#include file=userleft.asp--></td>
<td width="5">&nbsp;</td>
<td valign="top">
<table border="0" cellpadding="0" style="border-collapse: collapse"  width="760" class="ty1">
<%if sc<>"" then 
username=request.cookies("dick")("username")
id=trim(request("selectedid"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" & "history.back()" & "</script>"
end if
str2=split(id,",")
set rsk=server.createobject("adodb.recordset")
for i=0 to ubound(str2)
set rs0 = Server.CreateObject("ADODB.RecordSet")
sql0="select xxid from [DNJY_reply] where username='"&username&"' and id="&trim(str2(i))
rs0.open sql0,conn,1,1
if rs0.eof then
exit for                                      
end if
xxid=rs0("xxid")
Conn.Execute("Update DNJY_info Set hfcs=hfcs-1 where id="&xxid)
rs0.close
set rs0=nothing

sqlk="delete from [DNJY_reply] where id="&cstr(str2(i))
rsk.open sqlk,conn,1,3
next
response.write "<script language=JavaScript>" & chr(13) & "alert('ɾ���ظ��ɹ���');" & "window.location='user_hfgl.asp?"&C_ID&"'" & "</script>"
response.end
rsk.close
set rsk=nothing
Call CloseDB(conn)
else%>
<FORM name=thisForm action="user_hfgl.asp?<%=C_ID%>" method=POST>    

<%
dim ThisPage,Pagesize,Allrecord,Allpage
i=0
ii=1
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end if
username=request.cookies("dick")("username")
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_reply] where xxusername='"&username&"' order by hfsj "&DN_OrderType&"" 
rs.open sql,conn,1,1
if rs.eof then
response.write "<tr><td></td><td><li>û�м�¼��ظ�δ�����</td></tr>"
else
rs.Pagesize=20
Pagesize=rs.Pagesize
Allrecord=rs.Recordcount
Allpage=rs.Pagecount
if ThisPage<1 then                           
ThisPage=1
end if
On Error Resume Next
rs.move (ThisPage-1)*Pagesize
do while not rs.eof
b=trim(rs("b"))
bb=len(b)
id=rs("id")

Dim rsxx,sqlxx,biaoti
set rsxx=server.createobject("adodb.recordset")
sqlxx = "select * from [DNJY_info] where id="&cstr(rs("xxid"))
rsxx.open sqlxx,conn,1,1
if rsxx.eof then
errdick(0)
response.end
end If
biaoti=rsxx("biaoti")

%>
<tr>
<td width="30" height="25" align="center" class="ty2">
<p>���</td>
<td width="200" height="25" align="center" class="ty2">
�ظ�����</td>
<td width="150" height="25" align="center" class="ty2">
�����Ϣ</td>
<td width="50" height="25" align="center" class="ty2">
�ظ���</td>
<td width="70" height="25" align="center" class="ty2">
�ظ�ʱ��</td>
<td width="30" height="25" align="center" class="ty2">
���</td>
<td width="30" height="25" align="center" class="ty2">
<span style="text-decoration: none">ɾ��</span></td>
<td width="10" height="25" align="center" class="ty2">
��</td>
</tr>
<tr>
<td width="30" height="26" align="center" bordercolor="#E1E1E1" style="border-bottom-style: solid; border-bottom-width: 1">
<%=i+1%>
</td>
<td width="200" height="26"  bordercolor="#E1E1E1" style="border-bottom-style: solid; border-bottom-width: 1">
<%=left(rs("neirong"),20)%>
<td width="150" height="26"  bordercolor="#E1E1E1" style="border-bottom-style: solid; border-bottom-width: 1">
<a target="_blank" href=<%=x_path("",rsxx("id"),"",rsxx("Readinfo"))%>><%=left(biaoti,20)%><font color="#FF0000"></td>
<td width="50" height="26" align="center" bordercolor="#E1E1E1" style="border-bottom-style: solid; border-bottom-width: 1"><%if rs("username")<>"" then%><a target="_blank" href=preview.asp?username=<%=rs("username")%>&id=<%=rs("id")%>&<%=C_ID%>><%=rs("username")%><%else%>����<%end if%></td>
<td width="70" height="26" align="center" bordercolor="#E1E1E1" style="border-bottom-style: solid; border-bottom-width: 1"><%=datevalue(rs("hfsj"))%></td>
<td width="30" height="26" align="center" bordercolor="#E1E1E1" style="border-bottom-style: solid; border-bottom-width: 1"><%if rs("hfyz")=1 then%>��<%else%>��<%end if%></td>
<td width="30" height="26" align="center" bordercolor="#E1E1E1" style="border-bottom-style: solid; border-bottom-width: 1">
<font color="#CC5200"><a href="user_hfgl.asp?selectedid=<%=id%>&sc=ok&<%=C_ID%>"><font color="#FF0000">ɾ��</font></a></font></td>
<td width="10" height="26" align="center" bordercolor="#E1E1E1" style="border-bottom-style: solid; border-bottom-width: 1">
<input type="checkbox" name="selectedid" value="<%=id%>"></td>
</tr>
<%rsxx.close
set rsxx=nothing
i=i+1
rs.movenext
if i>=Pagesize then exit do
loop
end if
%>        <tr>
<td width="100%" height="25" colspan="8">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="49">
<tr> 
<td height="25" width="151">
<p align="center">
��</td>
<td height="25" width="126">
<p align="center">��</td>
<td height="25" width="118">
<p align="right">��</td>
<td height="25" width="200">
<p align="center">
<INPUT onclick=CheckAll(this.form) type=checkbox value=on name=chkall>ѡ�����м�¼ 
<input onclick=javascript:showoperatealert(1) type="submit" value="ɾ��" name="sc"></td>
</tr>
<tr> 
<td height="25" width="200" align="center" class="ty2">
����&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;����¼</td>
<td height="25" width="126" align="center" class="ty2">
�� <font color="#CC5200"><%=Allpage%></font> ҳ</td>
<td height="25" width="118" align="center" class="ty2">
�����ǵ� 
<font color="#CC5200"><%=ThisPage%></font> ҳ</td>
<td height="25" width="200" align="center" class="ty2">
<%
if ThisPage<2 then     
response.write "<font color=""#808080"">��ҳ</font>&nbsp;"
response.write "<font color=""#808080"">��һҳ</font>&nbsp;"     
else     
response.write "<a href=?page=1>��ҳ</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&">��һҳ</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">��һҳ</font>&nbsp;"
response.write "<font color=""#808080"">βҳ</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&">��һҳ</a>&nbsp;"   
response.write "<a href=?page="&Allpage&">βҳ</a>&nbsp;"     
end if
%></td>
</tr>
</table>
</td>
        <tr>
          <td width="99%" height="25" colspan="3"><p align="center"><font color="#0000FF">˵�����̴���ظ��Ѿ�����Ա��ˣ�</font><font color="#FF0000">���������ڵȴ�����Ա��ˣ�ǰ̨���ɼ���</font></td>
        </tr>
</tr>
</form>
<%end if%>
</table>
</td>
</tr>
</table>
</div>
<%Call CloseRs(rs)%>
<!--#include file=footer.asp-->
</body>
</html>
