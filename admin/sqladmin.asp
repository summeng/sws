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
<%Call OpenConn
Dim id,URL,Action,listnum,page,i,j,nn,n,filename,str2,ss,ipname
ss=request("ss")
ipname=Request("ipname")
Server.ScriptTimeout	=500		

URL						= Request.ServerVariables("URL")
Action					= Request("Action")


	If Action="del" Then
		Call Delip()
    ElseIf Action="BatchDelLog" Then
		Call BatchDelLog()    
	ElseIf Action="addip" Then
		Call addip()
	ElseIf Action="lock" Then
		Call lockIP()
	ElseIf Action="unlock" Then
		Call UnLockip()
	Else
		Call Main()
	End If


Sub Delip()
id=request("selectedid")
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('没有选择记录！');" &"window.location='sqladmin.asp'" & "</script>"
response.end
end if
str2=split(id,",")
set rs=server.createobject("adodb.recordset")
for i=0 to ubound(str2)
sql="delete from [DNJY_sqlin] where id="&cstr(str2(i))
rs.open sql,conn,1,3
next
response.write "<script language=JavaScript>" & chr(13) & "alert('删除成功！');" &"window.location='sqladmin.asp'" & "</script>"
response.end
Call CloseRs(rs)
Call Main()
End Sub

Sub addip()
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from DNJY_config"
rs.open sql,conn,1,3
if instr(rs("ip"),request("ip"))=0 Then
If rs("ip")="" Then
rs("ip")=request("ip")
else
rs("ip")=trim(rs("ip"))&"@"&request("ip")
End If
End If
rs.update
Call CloseRs(rs)
response.write "<script language=JavaScript>" & chr(13) & "alert('锁定IP:"&request("ip")&"成功！');" &"window.location='sqladmin.asp'" & "</script>"
End sub

Sub Lockip()
id = clng(request("id"))
set rs=server.createobject("adodb.recordset")
conn.execute("update DNJY_sqlin set Kill_ip="&DN_true&" where id="&id)
Call Main()
End sub

Sub UnLockip()
id = clng(request("id"))
conn.execute("update DNJY_sqlin set Kill_ip="&DN_False&" where id="&id)
Call Main()
End sub

sub BatchDelLog
	dim daynum,TotalCount
	daynum=request.form("daynum")
	'容错判断
	if daynum="" or not isnumeric(daynum) Then
	response.write "请输入正确参数！"
	Response.Write "<p align=""center""><< <a href=""javascript:history.go(-1)"">返回上一页</a></p>"
	response.end
	End if	
	
	if daynum<=1 Then
	response.write "1天内的数据不能删除！"
	Response.Write "<p align=""center""><< <a href=""javascript:history.go(-1)"">返回上一页</a></p>"
	response.end
	End if
	
	'开始删除	
	set rs=conn.execute("select count(*) from DNJY_sqlin where Kill_IP="&DN_False&" and SqlIn_TIME<"&DN_Now&"-"&daynum&"")
	TotalCount=rs(0)
	set rs=nothing
	if TotalCount>0 then
		conn.execute("delete from DNJY_sqlin where Kill_IP="&DN_False&" and SqlIn_TIME<"&DN_Now&"-"&daynum&"")	
		response.write "<script language=JavaScript>" & chr(13) & "alert('成功删除了"&daynum&"天前" & TotalCount &"条记录！');" &"window.location='?'" & "</script>"
	Else
	response.write "<script language=JavaScript>" & chr(13) & "alert('暂时没有符合删除条件的记录！');" &"window.location='?'" & "</script>"
	end if
	Call Main()
end Sub
%>

<%Sub Main()%>
  <style type="text/css">
<!--

table {
	font: 14px Tahoma, Verdana, "宋体";
}
a:link, a:visited {
	text-decoration: none;
	color: #036;
	font-family: Tahoma, Verdana, "宋体";
}
a:hover {
	text-decoration: none;
	color: #F90;
	font-family: Tahoma, Verdana, "宋体";
}
-->
</style>
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

//-->
    </SCRIPT>
<link href="../css/style.css" rel="stylesheet" type="text/css">

    <table width="98%" border="1" align="center" cellpadding="0" cellspacing="0">
<tr align=center bgcolor=#efefef>
<%set rs=server.createobject("adodb.recordset")
sql="select * from DNJY_sqlin"
If ss="1" Then sql=sql&" where SqlIn_IP='"&ipname&"'"
If ss="2" Then sql=sql&" where Kill_ip="&DN_True&""
If ss="3" Then sql=sql&" where Kill_ip="&DN_False&""
sql=sql&" order by id "&DN_OrderType&""           
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	response.write "暂无内容"
else
'分页的实现 
listnum=30
Rs.pagesize=listnum
page=Request("page")
if (page-Rs.pagecount) > 0 then
page=rs.pagecount
elseif page = "" or page < 1 then
page = 1
end if
rs.absolutepage=page
'编号的实现
j=rs.recordcount
j=j-(page-1)*listnum
i=0
nn=request("page")
if nn="" then
n=0
else
nn=nn-1
n=listnum*nn
end if%>
 <td width="5%" height=20>选择</td>
 <td width="5%" height=20>编号</td>
 <td width="15%"><font color=red>操作ＩＰ</font></td>
 <td width="5%">是否锁定</td>
 <td width="10%">操作页面</td>
 <td width="10%">操作时间</td>
 <td width="5%">提交方式</td>
 <td width="10%">提交参数</td>
 <td width="25%">提交数据</td>
 <td width="10%">操作</td>
</tr>
<FORM name=chkall action="?action=del&id=<%=rs("id")%>" method=POST>
<%do while not rs.eof and i<listnum
n=n+1%>
<tr align=center height=22>
<td> <INPUT type=checkbox value="<%=rs("id")%>" name=selectedid></td></td>
 <td><%=n%></td>
 <td><a href="http://www.ip138.com/ips.asp?ip=<%=rs("SqlIn_IP")%>" name="ipform" target="_blank"><%=rs("SqlIn_IP")%></a><a href="?ss=1&page=<%=Page%>&ipname=<%=rs("SqlIn_IP")%>"> 搜</a> <a href=<%=URL%>?action=addip&ip=<%=rs("SqlIn_IP")%>>锁</a></td>
<td><%	if rs("Kill_ip")=DN_true then 
			response.write "<font color='red'>已锁定</font>"
		else
			response.write "<font color='green'>已解锁</font>"
		end if
	%></td>
 <td><%=rs("SqlIn_WEB")%></td>
 <td><%=rs("SqlIn_TIME")%></td>
 <td><%=rs("SqlIn_FS")%></td>
 <td><%=rs("SqlIn_CS")%></td>
 <td><%=rs("SqlIn_SJ")%></td>
 <td><a href=<%=URL%>?action=del&selectedid=<%=rs("id")%>>删除</a>&nbsp;
 <%	if rs("Kill_ip")=DN_true then 
			response.write "<a href="&URL&"?action=unlock&id="&rs("id")&">解锁IP</a>"
		else
			response.write "<a href="&URL&"?action=lock&id="&rs("id")&">锁定IP</a>"
		end if
	%>
 
 </td>
</tr>
<%rs.movenext 
i=i+1 
j=j-1
loop%>
<tr>
<%filename=URL%>
<td colspan=10 align=right><%=Rs.recordcount%> 条记录&nbsp;&nbsp;<%=listnum%> 条记录/页&nbsp;&nbsp;共 <%=rs.pagecount%> 页 
      <% if page=1 then %>
      <%else%>
      <a href=<%=filename%>?ss=<%=ss%>&ipname=<%=ipname%>><strong>|<<</strong></a>
      <a href=<%=filename%>?page=<%=page-1%>&ss=<%=ss%>&ipname=<%=ipname%>><strong><<</strong></a>
      <a href=<%=filename%>?page=<%=page-1%>&ss=<%=ss%>&ipname=<%=ipname%>><b>[<%=page-1%>]</b></a>
      <%end if%><% if rs.pagecount=1 then%><%else%><b>[<%=page%>]</b><%end if%>
	  <% if rs.pagecount-page <> 0 then %>
      <a href=<%=filename%>?page=<%=page+1%>&ss=<%=ss%>&ipname=<%=ipname%>><b>[<%=page+1%>]</b></a>
      <a href=<%=filename%>?page=<%=page+1%>&ss=<%=ss%>&ipname=<%=ipname%>><strong>>></strong></a>
      <a href=<%=filename%>?page=<%=rs.pagecount%>&ss=<%=ss%>&ipname=<%=ipname%>><strong>>>|</strong></a>
	  <%end if%></td>
<%end if%></tr> 
</table>
<input onClick=CheckAll(this.form) type=checkbox value=on name=chkall>选中所有&nbsp;&nbsp;<input type=hidden name=action value=sendmail>
<input type=submit value="删除记录">
</FORM>

<form name="Log" method="post" action="?action=BatchDelLog">
        删除 
        <input name="daynum" type="text" value="5" size="3" maxlength="4">
        天前未锁定的记录(1天内的记录不会被删除) 
        <input type="submit" name="Submit" value="删除">
</form>

<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6699cc">
  <tr> 
    <td align="center" colspan="3">搜索查询</td>
  </tr>
  <tr bgcolor="#FFFFFF">
  <td><a href="?ss=2&page=<%=Page%>"><font size="2">已锁定</font></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="?ss=3&page=<%=Page%>"><font size="2">未锁定</font></a></td>
    <form name="form2" method="post" action="?ss=1&page=<%=Page%>">
      <td>
	  <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
          <tr bgcolor="#FFFFFF" align="center"> 
            <td width="35%"> <input name="ipname" type="text" id="selectkey" onFocus="this.value=''" value="请输入IP地址" size="30"> 
             <input type="submit" name="Submit" value="查 询" ></td>
          </tr>
        </table></td>
    </form>
  </tr>
</table>
<FONT COLOR="#ff0000">注:SQL注入访问限制改在“网站系统管理-字符IP限制设置”中进行管理,此处只提供注入信息管理,锁定与解锁无效!</FONT>
<TABLE width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
<TR>
	<TD width="10%"></TD>
	<TD align="center"> <A HREF="http://www.neeao.com" TARGET="_blank"><FONT COLOR="#000000">SQL</FONT></A><a href="http://www.ip126.com" TARGET="_blank"><FONT COLOR="#000000">通用防注入系统V3.0版</FONT></a></TD>
</TR>
</TABLE>
<%end Sub
Call CloseDB(conn)%>