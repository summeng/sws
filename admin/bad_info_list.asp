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
			thisForm.action="Bad_info_yz.asp";
			thisForm.submit();
		}
		}
		if (id==2)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="Bad_info_del.asp";
			thisForm.submit();
		}
		}
		if (id==3)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="Bad_info_yz.asp?dick=delyz";
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
if (confirm("真的要删除这条记录吗？删除后将无法恢复！"))
window.location = params ;
}
function test()
{
  if(!confirm('是否确定进行批量操作？操作后不能恢复！')) return false;
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
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">举 报 信 息 管 理</FONT></TD>
 </TR>
<tr>
<td height="31">
<FORM name=sou1 action="?dick=1" method=POST style="line-height: 150%; margin-top: 0; margin-bottom: 0">
<p align="center">按信息ID：<input type="text" name="T1" size="10">
<input type="submit" value="查找" name="B1">
</form>
</td>
<td>
<FORM name=sou3 action="?dick=3" method=POST style="line-height: 150%; margin-top: 0; margin-bottom: 0">
<p align="center">按关键词：<input type="text" name="T3" size="10">
<input type="submit" value="查找" name="B1">
</form>
</td>
</tr>
<tr>
<td width="864" colspan="2" height="28">
<p align="center"><a href="?dick=4">未(<font color="#FF0000">ⅹ</font>)验证举报</a>---<a href="?dick=5">已(√)验证的举报</a>---<a href="?dick="><font color="#FF0000">显示全部举报</font></a></td>
</tr>
</table>
</td>
</tr>
<FORM name=thisForm action="Bad_info_del.asp" method=POST>
<%function HTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")
    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")
    HTMLEncode = fString
end if
end function
dim k,dick
dim ThisPage,Pagesize,Allrecord,Allpage
k=1
Call OpenConn
set rs = Server.CreateObject("ADODB.RecordSet")
dick=trim(request("dick"))
Select Case dick
Case "1"
sql="select * from [DNJY_Bad_info] where infoid="&trim(request("T1"))&" order by addsj "&DN_OrderType&"" 
Case "3"
sql="select * from [DNJY_Bad_info] where Bad_info like '%"&trim(request("T3"))&"%' order by addsj "&DN_OrderType&"" 
Case "4"
sql="select * from [DNJY_Bad_info] where sh=0 order by addsj "&DN_OrderType&"" 
Case "5"
sql="select * from [DNJY_Bad_info] where sh=1 order by addsj "&DN_OrderType&""
Case Else
sql="select * from [DNJY_Bad_info] order by addsj "&DN_OrderType&"" 
End Select
rs.open sql,conn,1,1
if rs.eof then
response.write "<tr><td><li>没有记录"
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
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#000000">
<tr>
<td width="3%" height="25" align="center" bgcolor="#FEEEBC" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#999999">选择</td>
<td width="5%" height="25" align="center" bgcolor="#FEEEBC" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#999999">相关信息</td>
<td width="10%" bgcolor="#FEEEBC" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#999999">举报类型</td>
<td width="35%" bgcolor="#FEEEBC" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#999999">
<p align="center">举报补充内容</td>
<td width="13%"  bgcolor="#FEEEBC" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#999999">
<p align="center">举报日期</td>
<td width="10%" bgcolor="#FEEEBC" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#999999">
<p align="center">举报者IP</td>
<td width="10%" bgcolor="#FEEEBC" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#999999">举报者信箱</td>
<td width="3%" bgcolor="#FEEEBC" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#999999">
<p align="center">验证</td>
<td width="3%" bgcolor="#FEEEBC" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#999999" align="center">
操作</td>
</tr>

<%dim uid
do while not rs.eof
uid=rs("infoid")
        dim rs2,sql2,biaoti
        set rs2=server.createobject("adodb.recordset") 
		sql2="select * from DNJY_info where id="&rs("infoid")
		rs2.open sql2,conn,1,1
		biaoti=rs2("biaoti")
		rs2.close
		set rs2=nothing%>
<tr onmouseover="this.style.backgroundColor='#EFE8D6';return true;" onmouseout="this.style.backgroundColor='#FFFFFF';">
<td  height="25"  bgcolor="#cccccc"><INPUT type=checkbox value="<%=rs("id")%>" name=selectedid></td>
<td  height="25" align="center" bgcolor="#cccccc"><a target='_blank' href=../show.asp?id=<%=rs("infoid")%>><img alt='查看该回复相关信息,(信息ID:<%=UID%>)' border=0 src='images/goto.gif'></td>
<td  height="25" bgcolor="#cccccc"><%=rs("Bad_infolx")%></td>
<td  height="25" bgcolor="#cccccc"><%=HTMLEncode(rs("Bad_info"))%></td>
<td  height="25" bgcolor="#cccccc"><%=rs("addsj")%></td>
<td  height="25" bgcolor="#cccccc"><%=rs("ip")%></td>
<td  height="25" bgcolor="#cccccc"><a href="#" onClick="window.open('Bad_info_mail.asp?email=<%=rs("Bad_infomail")%>&biaoti=<%=biaoti%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=455,height=328,left=300,top=100')"><%=rs("Bad_infomail")%></a></td>
<td  height="25" bgcolor="#cccccc"><%if rs("sh")=1 then%><a href="Bad_info_yz.asp?dick=delyz&selectedid=<%=rs("id")%>">√</a><%else%><b><a href="Bad_info_yz.asp?selectedid=<%=rs("id")%>"><font color="#FF0000">ⅹ</font></a></b><%end if%></td>
<td  height="25" bgcolor="#cccccc"><font color="#FF0000">
<span id="followImg<%=k%>" style="CURSOR: hand" onclick="loadThreadFollow(<%=k%>,5)">
&nbsp;</span></font><a href="javascript:DoEmpty('Bad_info_del.asp?selectedid=<%=rs("id")%>&xid=<%=uid%>')">删除</a></td>
</tr>
<%
k=k+1
rs.movenext
if k>=Pagesize then exit do
loop
Call CloseRs(rs)
Call CloseDB(conn)
%>
</table>
<tr>
<td width="100%" height="1" colspan="6" style="border-top-style: solid; border-top-width: 1">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="1">
<tr> 
<td height="25" width="250" colspan="5">
<p align="left">
　<INPUT onclick=CheckAll(this.form) type=checkbox value=on name=chkall>选中所有<input onclick=javascript:showoperatealert(1) type="submit" value="批量验证" name="B1">
<input onclick=javascript:showoperatealert(3) type="submit" value="取消验证" name="B3">
<input name='submit1' type='submit' value=' 批量删除 ' onClick="return test();"  style="color: #FF0000">
</td> 

</tr>
<tr> 
<td height="25" width="800" align="right">
共有&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;条记录
共 <font color="#CC5200"><%=Allpage%></font> 页
现在是第 
<font color="#CC5200"><%=ThisPage%></font> 页
<%
if ThisPage<2 then     
response.write "<font color=""#808080"">首页</font>&nbsp;"
response.write "<font color=""#808080"">上一页</font>&nbsp;"     
else     
response.write "<a href=?page=1&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&dick&">首页</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&dick&">上一页</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">下一页</font>&nbsp;"
response.write "<font color=""#808080"">尾页</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&dick&">下一页</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&dick&">尾页</a>&nbsp;"     
end if
%></td>
</tr>

</table>
</td>
</tr>

</form>

</table>
</div>