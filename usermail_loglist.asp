<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->
<!--#include file=usercookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<script language="JavaScript" type="text/JavaScript">
function checkform()
{
  if (document.postart.searchkeys.value=="")
  {
    alert("请输入查询关键词！");
	document.postart.searchkeys.focus();
	return false;
  }
  if (document.postart.conditions.value==0)
  {
    alert("请选择查询条件！");
	document.postart.conditions.focus();
	return false;
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
			thisForm.action="?k=ok&lock=1&isse=0";
			thisForm.submit();
		}
		}
		if (id==2)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="?k=ok&lock=0&isse=0";
			thisForm.submit();
		}
		}

}

//-->
    </SCRIPT>
<script language=javascript>
//弹出和管理回复开始
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
if (confirm("真的要删除这条记录吗？删除后将无法恢复！"))
window.location = params ;
}
function test()
{
  if(!confirm('是否确定进行批量操作？操作后不能恢复！')) return false;
}
-->
</script>
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
<script language = "JavaScript">
function go(URL,cfmText)
{
	var ret;
	ret = confirm(cfmText);
	if(ret!=false)window.location=URL;
}
function CheckAll(form)
	{
	for (var i=0;i<form.elements.length;i++){
	var e = form.elements[i];
	if (e.name != 'chkall')
		e.checked = form.chkall.checked;
		}
	}
 </script>
 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>邮件日志管理</title>

</head>

<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
  <table width="980" border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" bgcolor="#FFFFFF" style="border-collapse: collapse">
    <tr>
      <td width="214" valign="top"><div align="center">
            <!--#include file=userleft.asp-->　
        </div></td>
    <td width="5">&nbsp;</td>
      <td width="760" align="right" valign="top" bgcolor="#FFFFFF">
	  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" class="ty1">
        <tr bgcolor="#FAFAFA">
          <td height="25"><span class="style5">我在线发送邮件的日志：</span></td>
		  </tr>        
          <tr>
            <td align="center">	
<%Call CloseRs(rs)

If conn.Execute("Select maillog From DNJY_mail")(0)>0 then'自动删除邮件日志
Dim rsdq
set rsdq=server.createobject("adodb.recordset")
sql="delete from [DNJY_log] where lock=0 and DateDiff("&DN_DatePart_D&",Sendtime,"&DN_Now&")>"&conn.Execute("Select maillog From DNJY_mail")(0)&""
rsdq.open sql,conn,1,3
End if	
Dim isse,conditions,searchkeys,pp,caozuocs,page,pages,j,submitname,pagecanshu,DelID,str2
isse=ReplaceUnsafe(request.QueryString("isse"))
conditions=ReplaceUnsafe(Request.Form("conditions"))
if conditions="" then conditions=ReplaceUnsafe(Request.QueryString("conditions"))
searchkeys=ReplaceUnsafe(Request.Form("searchkeys"))
if searchkeys="" then searchkeys=ReplaceUnsafe(Request.QueryString("searchkeys"))

If request("key")="ok" And request("DelID")="" And request("k")<>"ok" Then'批量删除
Response.Write ("<script language=javascript>alert('没有选择记录!');history.go(-1);</script>")
  response.end
end If
If request("key")="ok" And request("DelID")<>"" And request("k")<>"ok" Then
id=trim(request("DelID"))
str2=split(id,",")
Set rs=Server.CreateObject("ADODB.RecordSet") 
for i=0 to ubound(str2)'循环开始
sql="delete from [DNJY_log] where username='"&username&"' and lock=0 and id="&cstr(str2(i))
rs.open sql,conn,1,3
Next'循环结束
response.write "<script language=JavaScript>" & chr(13) & "alert('操作成功！锁定邮件不能删除。');" &"window.location='?isse=0'" & "</script>"
response.end
End If

If request("k")="ok" And request("DelID")="" And request("lock")<>"" Then'批量加解锁
Response.Write ("<script language=javascript>alert('没有选择记录!');history.go(-1);</script>")
  response.end
end If
If request("k")="ok" And request("DelID")<>"" And request("lock")<>"" Then
Dim rslog
id=trim(request("DelID"))
str2=split(id,",")
for i=0 to ubound(str2)'循环开始
set rslog=server.createobject("adodb.recordset")
sql="select * from DNJY_log where username='"&username&"' and id="&cstr(str2(i))
rslog.open sql,conn,1,3
If request("lock")=1 Then
rslog("lock")=1
Else
rslog("lock")=0
End if
rslog.update
rslog.close
set rslog=nothing
Next'循环结束
response.write "<script language=JavaScript>" & chr(13) & "alert('操作成功！');" &"window.location='?isse=0'" & "</script>"
response.end
End If

If request("lock")="shno"  Then Conn.Execute("Update DNJY_log Set lock=0 where id="&cstr(Request("id")))
If request("lock")="sh"  Then Conn.Execute("Update DNJY_log Set lock=1 where id="&cstr(Request("id")))

if isse=0 Then
set rs=server.createobject("adodb.recordset")
sql="select * from DNJY_log where username='"&username&"' order by id desc"
end if

if isse=1 Then
	if conditions=1 then		
	sql="select * from DNJY_log where  username='"&username&"' and outbox='"&searchkeys&"' order by id"
	end If
	if conditions=2 then		
	sql="select * from DNJY_log where  username='"&username&"' and inbox='"&searchkeys&"' order by id"
	end If
	if conditions=3 then		
	sql="select * from DNJY_log where  username='"&username&"' and title='"&searchkeys&"' or content='"&searchkeys&"' order by id"
	end if		
	if conditions=4 then		
	sql="select * from DNJY_log where  username='"&username&"' and lock=1 order by id"
	end if	
end if
%>

<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td valign="top"><form name="postart" method="post" action="usermail_loglist.asp?isse=1" onSubmit="return checkform()">
              <br>
              查找信息：
	
                          <select name="conditions">
                            <option value="0" selected>查找条件</option>                            
							<option  value="1">发送信箱</option>
							<option  value="2">接收信箱</option>
							<option  value="3">标题或内容</option>
							<option  value="4">锁定邮件</option>
                          </select>
              <div  style="display:inline" id="gjz">关键字&nbsp;：</div>          
			 <input id="searchkeys" type="text"name="searchkeys" size="40" value="" tips="请输入关键字"/>
             <input style="height:20px;" name="Submit" type="submit" class="button" value="提交">
            </form>            </td>
          </tr>          
        </table>
	  <table width="90%" border="0" cellspacing="1" align="center" class="list">
    <tr>
      <td><font color="#ff0000">注意：未锁定的邮件日志将在 <%=conn.Execute("Select maillog From DNJY_mail")(0)%> 天内自动删除！<br>为节省资源，无用邮件请及时删除。网站管理员有权删除邮件（包括锁定的），并不一定通知会员。重要资料自行保存。</font></TD>
  </TR>
</table>
 <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="760">
  <form name="thisForm" action="?key=ok" method="POST">
    <TR> 
    <TD height="5" colspan="10" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">邮件日志管理</FONT></TD>
  </TR>
	
    <tr>
      <td width="10" >&nbsp;</td>
      <td width="20" ><div align="center">ID号</div></td>
	  <td width="50" ><div align="center">操作人员</div></td>
      <td width="25"><div align="center">发信邮箱</div></td>
      <td width="25" ><div align="center">收信邮箱</div></td>
	  <td width="170" ><div align="center">邮件标题</div></td>
      <td width="270" ><div align="center">邮件内容</div></td>
	  <td width="150" ><div align="center">发送时间</div></td>
      <td width="25" ><div align="center">是否锁定</div></td>
      <td width="25"><div align="center">操作</div></td>
    </tr>
<%
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1
if rs.eof and rs.bof then
    response.Write("<tr><td colspan='8'><div align='center'>没有记录</div> </td></tr>")
else
	rs.PageSize=20
	page=clng(request("page"))
	if page=0 then page=1 
	pages=rs.pagecount
	if page > pages then page=pages
	rs.AbsolutePage=page 
	pp=1
	for j=1 to rs.PageSize 	
	if isse=0 then
	caozuocs="isse=0&id="&rs("id")&"&page="&page
	end if
	if isse=1 then
			caozuocs="isse=1&id="&rs("id")&"&page="&page&"&conditions="&conditions&"&searchkeys="&searchkeys	 
	end if
%>

    <tr>
      <td width="10"><input name="DelID" type="checkbox" id="DelID" value="<%=rs("id")%>">
	  <input type="hidden" name="key" value="ok"></td>
	  <td width="20"><div align="center"><%=rs("id")%></div></td>
	  <td width="50"><div align="center"><%=rs("username")%></div></td>      
      <td width="25"><%=rs("outbox")%></td>
	  <td width="25"><%=rs("inbox")%></td>
	  <td width="170"><%=rs("title")%></td>
	  <td width="270"><%=rs("content")%></td>
	  <td width="150"><%=rs("Sendtime")%></td>
      <td width="25">
	    <div align="center">
	      <%
		if rs("lock")=1 then 
			response.Write "<a href=?lock=shno&id="&rs("id")&"&isse=0><font color=""#339933"">是</font></a>" 
		else
			response.Write "<a href=?lock=sh&id="&rs("id")&"&isse=0><font color=""#CC0000"">否</font></a>" 
		end if
		%>
        </div></td>
      <td width="25"><div align="center"><a href="javascript:go('?key=ok&DelID=<%=rs("id")%>','您确定要删除？')">删除</a>  </div></td>
    </tr>
<%
rs.movenext
pp=pp+1
if rs.eof then exit for
next%>	
	
    
	
    <tr>
      <td colspan="13" > <br> <div align="left">        
          <label>
            <input onclick=CheckAll(this.form) type="checkbox" name="chkall" value="checkbox">
            全选</label>		           
		  <input onclick=javascript:showoperatealert(1) type="submit" value="批量锁定" name="B1">
          <input onclick=javascript:showoperatealert(2) type="submit" value="批量解锁" name="B2">	  
          <input name='submit1' type='submit' value=' 批量删除 ' onClick="return test();"  style="color: #FF0000">        
      </div>
	  </form> 
<%
 submitname="usermail_loglist.asp"
 if isse=0 then
 pagecanshu="isse=0"
 end if
 if isse=1 then
	 if conditions=1 then
	 	pagecanshu="isse=1&conditions="&conditions&"&searchkeys="&searchkeys
	 end if
	 if conditions=2 then
	 	pagecanshu="isse=1&conditions="&conditions&"&searchkeys="&searchkeys
	 end If
	 if conditions=3 then
	 	pagecanshu="isse=1&conditions="&conditions&"&searchkeys="&searchkeys
	 end if	 
 end if
 
 %>
 <div align="center">
 <form method=Post action="<%=submitname&"?"&pagecanshu%>">
 <%if Page<2 then      
	response.write "首页 上一页&nbsp;"
 else
    response.write "<a href="&submitname&"?"&pagecanshu&"&page=1>首页</a>&nbsp;"
    response.write "<a href="&submitname&"?"&pagecanshu&"&page="&(Page-1)&">上一页</a>&nbsp;"
 end if
 if rs.pagecount-page<1 then
    response.write "下一页 尾页"
 else
    response.write "<a href="&submitname&"?"&pagecanshu&"&page="&(page+1)&">"
    response.write "下一页</a> <a href="&submitname&"?"&pagecanshu&"&page="&rs.pagecount&">尾页</a>"
  end if
   response.write "&nbsp;页次：<strong><font color=red>"&Page&"</font>/"&rs.pagecount&"</strong>页 "
   response.write "&nbsp;共<b><font color='#FF0000'>"&rs.recordcount&"</font></b>条记录 <b>"&rs.pagesize&"</b>条记录/页"
   response.write " 转到：<input type='text' name='page' size=4 maxlength=10  value="&page&">"
   response.write " <input  type='submit'  class='button' value=' Goto '  name='cndok'></span>"     
%>
 </form>
 </div>
      </td>      
    </tr>
<%
end if
set rs=Nothing
%>		
		
		
		</td>
                </tr>
              </table>
                <p> </td>
          </tr>
        </form>
      </table>	  
	  </td>
      <td width="4" align="right" valign="top" bgcolor="#e6e6e6"></td>
    </tr>
    <tr>
      <td width="980" height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>
