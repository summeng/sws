<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=usercookies.asp-->
<!--#include file=top.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%dim str2
Call OpenConn
c=trim(request("c"))
if trim(request("linkdel"))="1" then
id=trim(request("selectedid"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('没有选择记录！');" & "history.back()" & "</script>"
response.end
end If

str2=split(id,",")
set rs=server.createobject("adodb.recordset")
for i=0 to ubound(str2)
sql="delete from [DNJY_comlink]  where id="&cstr(str2(i))
rs.open sql,conn,1,3
Next
response.Write "<script language=javascript>alert('删除链接成功!');location.replace('user_link.asp?c="&c&"');</script>"
response.end
Call CloseRs(rs)
else

dim add,comgg
add=Request.form("add")
username=request.cookies("dick")("username")
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_user] where username='"&username&"'" 
rs.open sql,conn,1,1
if rs.eof then
response.write "<br>"
response.write "<li>参数错误！"
response.end
end if
vip=rs("vip")
If rs("sdkg")=0 Then
response.write "<br><br><div align='center'>您未开有店铺或店铺未经审核，不能添加友情链接！<a href='user.asp'>返回</a></div>"
Call CloseRs(rs)
response.end
End if
%>
<SCRIPT language=javascript>
<!--
function CheckForm()
{
if(document.thisForm1.web.value.length<1)
	{
	    alert("网站名称不能为空!");
	    return false;
	}
	if(document.thisForm1.url.value.length<1)
	{
	    alert("网站地址不能为空!");
	    return false;
	}
			if(document.thisForm1.webabout.value.length<1)
	{
	    alert("网站简介不能为空!");
	    return false;
	}
	


}

//-->
</SCRIPT>
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
			thisForm.action="user_link.asp?linkdel=1&"&C_ID&"";
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
//-->
</SCRIPT>
<body topmargin="0">
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="980" height="426">
    <tr>
      <td width="214" height="354" valign="top"  bgcolor="#FFFFFF"><div align="center">
      
        <!--#include file=userleft.asp--></div></td>
		<td width="5">&nbsp;</td>
<td valign="top" >
<%if add="" then %>
<table width="760" border="0" align="right" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse" class="ty1">
<tr>
 <td height="25" class="ty2" colspan="2" align="center"><span class="style5">友情链接仅在网络店铺有效</span></td>
</tr>
<%Dim rsnews
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
sql="select id from DNJY_comlink where username='"&username&"'"
rsnews.Open sql,conn,1,1
If userlink>0 And rsnews.recordcount>=userlink Then
response.write "<tr><td><div align='center'>您添加的链接已达到限定数量，不能再添加！</div></td></tr>"
else
%>
<tr>
<td width="18%">　</td>
<td width="82%">

<form method="POST" name="thisForm1" action="user_link.asp?<%=C_ID%>" style="line-height: 100%; margin-top: 0; margin-bottom: 0">
<p style="line-height: 220%; margin-top: 0; margin-bottom: 0">
网站/店铺名称:<input type="text" name="web" size="30" maxlength="10"> *</p>
<p style="line-height: 220%; margin-top: 0; margin-bottom: 0">
网站/店铺地址:<input type="text" name="url" size="30" value="http://"> *</p>
<p style="line-height: 220%; margin-top: 0; margin-bottom: 0">
网站/店铺简介:<input type="text" name="webabout" size="36" maxlength="30"> *</p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input border="0" onclick="javascript:return CheckForm();" type="submit" value="添加新链接" name="add">
</form></td>
</tr>
<%
rsnews.close
set rsnews=Nothing
End If%>
<tr>
<td colspan="2">
<table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" height="1" bordercolor="#808000" valign="top">
 <FORM name=thisForm action="user_link.asp?linkdel=1&<%=C_ID%>" method=POST>

<%
dim k,cnmai
dim ThisPage,Pagesize,Allrecord,Allpage
k=1
set rs = Server.CreateObject("ADODB.RecordSet")
cnmai=trim(request("cnmai"))
sql="select * from [DNJY_comlink] where username='"&username&"' order by jrsj "&DN_OrderType&"" 
rs.open sql,conn,1,1
if rs.eof then
response.write "<li>没有链接记录"
end if
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end if
rs.Pagesize=10
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
<td width="8%" height="24" align="center" class="ty2">选择</td>
<td width="27%" height="24" class="ty2">
网站名称</td>
<td width="29%" height="24" class="ty2">
<p align="center">网站地址</td>
<td width="32%" height="24" class="ty2">　</td>
</tr>
<%
dim uid
do while not rs.eof
uid=rs("id")
%>

<tr>
<td width="8%" height="24" align="center" style="border-bottom:1px solid #CCCCCC; border-top-width: 1px; border-left-width:1px; border-right-width:1px">
<INPUT type=checkbox value="<%=uid%>" name=selectedid></td>
<td width="27%" height="24" style="border-bottom:1px solid #CCCCCC; border-top-width: 1px; border-left-width:1px; border-right-width:1px"><a target="_blank" href=<%=rs("url")%>><%=rs("web")%></a></td>
<td width="29%" height="24" align="center" style="border-bottom:1px solid #CCCCCC; border-top-width: 1px; border-left-width:1px; border-right-width:1px"><a target="_blank" href=<%=rs("url")%>><%=rs("url")%></td>
<td width="32%" height="24" style="border-bottom:1px solid #CCCCCC; border-top-width: 1px; border-left-width:1px; border-right-width:1px">
<p align="center"><a target="_blank" href="user_link_e.asp?selectedid=<%=uid%>&<%=C_ID%>">修改</a> | <a href="user_link.asp?selectedid=<%=uid%>&linkdel=1&<%=C_ID%>">删除该链接</a></td>
</tr>
<%
k=k+1
rs.movenext
if k>=Pagesize then exit do
loop
Call CloseRs(rs)
%>
<tr>
<td width="100%" height="1" colspan="4">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="1">
<tr> 
<td height="25" width="395" colspan="3">
<p align="left">
 　<INPUT onclick=CheckAll(this.form) type=checkbox value=on name=chkall>选中所有<input type="submit" value="删除" name="B2"></td>
<td height="25" width="200">
<p align="center">
　</td>
</tr>
<tr> 
<td height="25" width="151" align="center">
共有&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;条记录</td>
<td height="25" width="126" align="center">
共 <font color="#CC5200"><%=Allpage%></font> 页</td>
<td height="25" width="118" align="center">
现在是第 
<font color="#CC5200"><%=ThisPage%></font> 页</td>
<td height="25" width="200" align="center">
<%
if ThisPage<2 then     
response.write "<font color=""#808080"">首页</font>&nbsp;"
response.write "<font color=""#808080"">上一页</font>&nbsp;"     
else     
response.write "<a href=?page=1&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&cnmai="&cnmai&"&"&C_ID&">首页</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&cnmai="&cnmai&"&"&C_ID&">上一页</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">下一页</font>&nbsp;"
response.write "<font color=""#808080"">尾页</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&cnmai="&cnmai&"&"&C_ID&">下一页</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&cnmai="&cnmai&"&"&C_ID&">尾页</a>&nbsp;"     
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
<%else 
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_comlink"
rs.open sql,conn,1,3
rs.addnew
rs("username")=username
rs("web")=trim(request("web"))
rs("url")=trim(request("url"))
rs("webabout")=trim(request("webabout"))
rs("jrsj")=now()
rs.update
Call CloseRs(rs)

response.write "<li>恭喜你，链接添加成功！！"
response.write "<meta http-equiv=refresh content=""1;URL=user_link.asp?"&C_ID&""">"
end if%> 
</form>
</table></td>
</tr>
</table>
</td>
</tr>
</table>
</div><!--#include file=footer.asp-->
</body>
</html><%end If%>