<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
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
<title>asp验证码后台管理系统</title>
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
if (confirm("真的要删除这条记录吗？删除后将无法恢复！"))
window.location = params ;
}
function test()
{
  if(!confirm('是否确定进行批量操作？操作后不能恢复！')) return false;
}
-->
</script>
</head>

<body>
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#cccccc">
  <tr>
    <td><a href="admin_check_add.asp">添加问题</a></td>
    <td height="13"><div align="center">验证问答管理<font color="#FF0000">（建议不要使用系统默认的问答，自行修改或添加并经常更换）</font></div></td>
    
  </tr>
</table>

<table width="95%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#000000">
  <tr> 
    <td width="5%" height="25" align="center" bgcolor="#CCCCCC">ID</td>
    <td width="43%" align="center" bgcolor="#CCCCCC">问题</td>
    <td width="29%" align="center" bgcolor="#CCCCCC">答案</td>
    <td width="23%" align="center" bgcolor="#CCCCCC">操作</td>
  </tr>
  <%
Dim totaluser,allPages,n,keywords
if rs.eof and rs.bof then
response.Write("没有记录")
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
    <td align="center"><a href="admin_check_edit.asp?id=<%=rs("id")%>">修改</a> <a href="javascript:DoEmpty('admin_check_delete.asp?id=<%=rs("id")%>')">删除</a>  <%if rs("xs")=1 then 
   response.write "<a href=admin_check_diaoyong.asp?action=rg&id="&rs("id")&"><font color='#0000ff'><font color='#0000ff'>正常显示</font></a>"
   else   
  response.write "<a href=admin_check_diaoyong.asp?action=wg&id="&rs("id")&"><font color='#FF0000'>已经隐藏</font></a>" 
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
    <td>　</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 

      <td height="30" align="center"> 
<%
call listPages()
end if

sub listPages() 
'分页
'if allpages<=1 then exit sub
response.write "<br>总记录数<font color=red>"&totaluser&"</font> &nbsp; 每页<font color=red>20</font>&nbsp; "
Response.Write "<font class='contents'> 页次：</font><font class='contents'>"&Page&"</font><font class='contents'>/"&n&"页</font> "  
if page = 1 then
response.write "<font color=darkgray>首页 前页</font>"
else
response.write "<a href=admin_check.asp?keywords="&keywords&"&page=1>首页</a> <a href=admin_check.asp?keywords="&keywords&"&page="&page-1&">前页</a>"
end if
if page = allpages then
response.write "<font color=darkgray> 下页 末页</font>"
else
response.write " <a href=admin_check.asp?keywords="&keywords&"&page="&page+1&">下页</a> <a href=admin_check.asp?keywords="&keywords&"&page="&allpages&">末页</a>"
end if
end sub%>
      </td>
  </tr>
</table>
<%Call CloseRs(rs)
Call CloseDB(conn)%>

</body>
</html>
