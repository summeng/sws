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
call checkmanage("06")%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%Call OpenConn
If request("yz")="yes" Or request("yz")="no" Then'��ʾ������
Dim str2,i
Server.ScriptTimeout=50000
id=trim(request("DelID"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" &"window.location='admin_announce.asp'" & "</script>"
response.end
end if
str2=split(id,",")
for i=0 to ubound(str2)
set rs=server.createobject("adodb.recordset")
sql="select ifshow from [DNJY_announce] where id="&trim(str2(i))
rs.open sql,conn,1,3
If request("yz")="yes" Then
rs("ifshow")=1
Else
rs("ifshow")=0
End if
rs.update
Next
Call CloseRs(rs)
End If
If request("tj")="yes" Or request("tj")="no" Then'�Ƽ����
id=trim(request("ID"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" &"window.location='admin_announce.asp'" & "</script>"
response.end
end if
set rs=server.createobject("adodb.recordset")
sql="select tuijian from [DNJY_announce] where id="&trim(id)
rs.open sql,conn,1,3
If request("tj")="yes" Then
rs("tuijian")=1
Else
rs("tuijian")=0
End if
rs.update
Call CloseRs(rs)
End If
If request("gd")="yes" Or request("gd")="no" Then'�̶����
id=trim(request("ID"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" &"window.location='admin_announce.asp'" & "</script>"
response.end
end if
set rs=server.createobject("adodb.recordset")
sql="select newstop from [DNJY_announce] where id="&trim(id)
rs.open sql,conn,1,3
If request("gd")="yes" Then
rs("newstop")=1
Else
rs("newstop")=0
End if
rs.update
Call CloseRs(rs)
End If%>
<title>��������</title>
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<script language=javascript src=../Include/mouse_on_title.js></script>
<script language=javascript>
<!--
function left_menu(meval)
{
  var left_n=eval(meval);
  if (left_n.style.display=="none")
  { eval(meval+".style.display='';"); }
  else
  { eval(meval+".style.display='none';"); }
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
-->
</script>
<script language="JavaScript">
<!--
function CheckAll(form)
{
for (var i=0;i<form.elements.length;i++)
{
var e = form.elements[i];
if (e.name != 'chkall')
e.checked = form.chkall.checked;
}
}

//-->
</script>
<SCRIPT language=JavaScript>
function showoperatealert(id)
{
		if (id==1)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="admin_announce.asp?yz=yes";
			thisForm.submit();
		}
		}
		if (id==2)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="admin_announce.asp?yz=no";
			thisForm.submit();
		}
		}
		if (id==3)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="admin_announce_modi.asp?del=yes";
			thisForm.submit();
		}
		}
}

//-->
    </SCRIPT>
</head>
<BODY background="img/back.gif"><br>

<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#000000">
<tr class=backs><td align="center" class=td height=18 colspan="22">��վ�������&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_announce_modi.asp?Result=Add"><font color="#FF0000">��ӹ���</font></a></td>
  </tr>
 <FORM name="thisForm" action="admin_announce_modi.asp?del=yes" method="POST">
  <tr> 
    <td width="5%" height="25" align="center" bgcolor="#cccccc">ѡ��</td>
    <td width="30%" align="center" bgcolor="#cccccc">�������</td>
    <td width="20%" align="center" bgcolor="#cccccc">��������</td>
    <td width="15%" align="center" bgcolor="#cccccc">����ʱ��</td>
    <td width="15%" align="center" bgcolor="#cccccc">����</td>
  </tr>
<%Dim city_oneid,city_twoid,city_threeid,selectm,selectkey,selectid,ProdId,comment,id,page,pages,j,photosl
page=clng(request("page"))
city_oneid=trim(request("city_oneid"))
city_twoid=trim(request("city_twoid"))
city_threeid=trim(request("city_threeid"))
If city_oneid="" then city_oneid=0
If city_twoid="" then city_twoid=0
If city_threeid="" then city_threeid=0
selectkey=trim(request.form("selectkey"))
selectm=trim(request.form("selectm"))
if selectkey="" then
selectkey=request.QueryString("selectkey")
end if
if selectm="" then
selectm=request.QueryString("selectm")
end if

set rs=server.createobject("adodb.recordset")
sql="select * from DNJY_announce  "

			select case selectm
			case ""
            sql=sql
		    case "name"
			sql=sql&" where title like '%"&selectkey&"%'"
			case "ProdId"
            If  Not isNumeric(selectkey) Then  
            response.write "���ӦΪ���֣�"     
            response.End
            else
			sql=sql&" where Id=cint("&selectkey&")"
            End if
			case "content"
			sql=sql&"where content like '%"&selectkey&"%'"			
		    end select

sql=sql&" order by id "&DN_OrderType&""
rs.open sql,conn,1,1
if rs.eof and rs.bof then
%>
<tr bgcolor="#FFFFFF"><td colspan="22" align="center">��ʱû����ؼ�¼</td>
</tr>
<%response.end
Else
rs.pagesize=15
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.absolutepage=page  
    
 for j=1 to rs.pagesize%>
  <tr bgcolor="#ffffff"> 
    <td height="22" align="center"><input name="DelID" type="checkbox" id="DelID" value="<%=rs("id")%>"></td>	
    <td><%If Rs("Solidtop")=1 Then response.write "<img src=""../img/top4.gif"" width=""15"" height=""13"" border=0 alt=""�̶�"">"%><a href="#" ONCLICK="window.open('../announce_show.asp?id=<%=rs("id")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=500,height=450,left=300,top=20')"><%=rs("title")%></a> </td>
    <TD align="center"><%=rs("city_one")%><%If rs("city_two")<>"" Then response.write "/"&rs("city_two")%><%If rs("city_three")<>"" Then response.write "/"&rs("city_three")%></TD>
   <td><%=rs("addtime")%></a></td>
    <td align="center"><%if rs("activity")=1 then%>��ʾ<%else%><font color=red>����</font><%end if%> <a href="admin_announce_modi.asp?Result=Modify&id=<%=rs("id")%>">�޸�</a>
	<a href="javascript:DoEmpty('admin_announce_modi.asp?DelID=<%=rs("id")%>&del=yes');">ɾ��</a></td>
 
<%
rs.movenext
if rs.eof then exit For
next%>
</table>
  <p align="center">
    <input type="checkbox" name="chkall" onClick="CheckAll(this.form)">
    ȫ��ѡ��
<input onclick=javascript:showoperatealert(1) type="submit" value="������ʾ" name="B1">
<input onclick=javascript:showoperatealert(2) type="submit" value="��������" name="B2">	  
<input name='submit1' type='submit' value=' ����ɾ�� ' onClick="return test();"  style="color: #FF0000">	
  </p>
</form>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr bgcolor="#ffffff"> 
<form method=post action="admin_announce.asp">  
      <td height="30" align="center"> 
    <%if page<2 then      
    response.write "��ҳ ��һҳ&nbsp;"
  else
    response.write "<a href=admin_announce.asp?page=1&id="&trim(request("id"))&"&fy=ok&selectm="&selectm&"&selectkey="&selectkey&">��ҳ</a>&nbsp;"
    response.write "<a href=admin_announce.asp?page=" & page-1 & "&id="&trim(request("id"))&"&fy=ok&selectm="&selectm&"&selectkey="&selectkey&">��һҳ</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "��һҳ βҳ"
  else
    response.write "<a href=admin_announce.asp?page=" & (page+1) & "&id="&trim(request("id"))&"&fy=ok&selectm="&selectm&"&selectkey="&selectkey&">"
    response.write "��һҳ</a> <a href=admin_announce.asp?page="&rs.pagecount&"&id="&trim(request("id"))&"&fy=ok&selectm="&selectm&"&selectkey="&selectkey&">βҳ</a>"
  end if
   response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&page&"</font>/"&rs.pagecount&"</strong>ҳ "
    response.write "&nbsp;��<b><font color='#ff0000'>"&rs.recordcount&"</font></b>����¼ <b>"&rs.pagesize&"</b>����¼/ҳ"
   response.write " ת����<input type='text' name='page' size=4 maxlength=10 class=input value="&page&">"
   response.write " <input class=input type='submit'  value=' goto '  name='cndok'></span></p>"     
%>
      </td></form>
  </tr>
</table>
<%end if
Call CloseRs(rs)
Call CloseDB(conn)%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6699cc">
  <tr> 
    <td align="center">������ѯ</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <form name="form2" method="post" action="admin_announce.asp">
      <td>
	  <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
          <tr bgcolor="#FFFFFF" align="center"> 
            <td > <input name="selectkey" type="text" id="selectkey" onFocus="this.value=''" value="������ؽ���" size="30"> 
             <select name="selectm" id="select">			   
                <OPTION VALUE="name">������</OPTION>                
                <OPTION VALUE="content">������</OPTION>
				<OPTION VALUE="ProdId">�����</OPTION> 
              </select><input type="submit" name="Submit" value="�� ѯ" ></td>
          </tr>
        </table></td>
    </form>
  </tr>
</table>

</body>
</html>
