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
<%call checkmanage("14")
Call OpenConn
dim Page,pages,j	
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from DNJY_wenda order by id asc"
	rs.open sql,conn,1,2
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>asp��֤���̨����ϵͳ</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	background-color: #BDBEDE;
}
.style1 {color: #FF0000}
-->
</style>
<script language=javascript>
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
</head>

<body>
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#cccccc">
  <tr>
    <td><a href="admin_check_add.asp">�������</a></td>
    <td height="13"><div align="center">��֤�ʴ����<font color="#FF0000">�����鲻Ҫʹ��ϵͳĬ�ϵ��ʴ������޸Ļ���Ӳ�����������</font></div></td>
    
  </tr>
</table>

<table width="95%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#000000">
  <tr> 
    <td width="5%" height="25" align="center" bgcolor="#CCCCCC">ID</td>
    <td width="43%" align="center" bgcolor="#CCCCCC">����</td>
    <td width="29%" align="center" bgcolor="#CCCCCC">��</td>
    <td width="23%" align="center" bgcolor="#CCCCCC">����</td>
  </tr>
  <%
Dim totaluser,allPages,n,keywords
if rs.eof and rs.bof then
response.Write("û�м�¼")
Else
pages = 20
totaluser=RS.RecordCount
rs.pageSize = pages
allPages = rs.pageCount
page = Request("page")
If not isNumeric(page) then page=1
if isEmpty(page) or clng(page) < 1 then
page = 1
elseif clng(page) >= allPages then
page = allPages 
end if
				If totaluser Mod pages=0 Then  
					n= totaluser \ pages  
				Else
					n= totaluser \ pages+1  
				End If
rs.AbsolutePage = page
    
Do While Not rs.eof and pages>0 
%>
  <tr bgcolor="#FFFFFF"> 
    <td height="22" align="center"><%=rs("id")%></td>
    <td><%=rs("wenti")%></td>
    <td><%=rs("daan")%></td>
    <td align="center"><a href="admin_check_edit.asp?id=<%=rs("id")%>">�޸�</a> <a href="javascript:DoEmpty('admin_check_delete.asp?id=<%=rs("id")%>')">ɾ��</a>  <%if rs("xs")=1 then 
   response.write "<a href=admin_check_diaoyong.asp?action=rg&id="&rs("id")&"><font color='#0000ff'><font color='#0000ff'>������ʾ</font></a>"
   else   
  response.write "<a href=admin_check_diaoyong.asp?action=wg&id="&rs("id")&"><font color='#FF0000'>�Ѿ�����</font></a>" 
   end if%></td>
  </tr>
  <%
pages = pages - 1
rs.movenext
if rs.eof then exit do
loop                                                    
%>
</table>
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>��</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 

      <td height="30" align="center"> 
<%
call listPages()
end if

sub listPages() 
'��ҳ
'if allpages<=1 then exit sub
response.write "<br>�ܼ�¼��<font color=red>"&totaluser&"</font> &nbsp; ÿҳ<font color=red>20</font>&nbsp; "
Response.Write "<font class='contents'> ҳ�Σ�</font><font class='contents'>"&Page&"</font><font class='contents'>/"&n&"ҳ</font> "  
if page = 1 then
response.write "<font color=darkgray>��ҳ ǰҳ</font>"
else
response.write "<a href=admin_check.asp?keywords="&keywords&"&page=1>��ҳ</a> <a href=admin_check.asp?keywords="&keywords&"&page="&page-1&">ǰҳ</a>"
end if
if page = allpages then
response.write "<font color=darkgray> ��ҳ ĩҳ</font>"
else
response.write " <a href=admin_check.asp?keywords="&keywords&"&page="&page+1&">��ҳ</a> <a href=admin_check.asp?keywords="&keywords&"&page="&allpages&">ĩҳ</a>"
end if
end sub%>
      </td>
  </tr>
</table>
<%Call CloseRs(rs)
Call CloseDB(conn)%>

</body>
</html>
