<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
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
response.write "<script language=JavaScript>" & chr(13) & "alert('没有选择记录！');" &"window.location='user_Temporary.asp'" & "</script>"
response.end
end if
str2=split(id,",")
set rs=server.createobject("adodb.recordset")
for i=0 to ubound(str2)
sql="delete from [DNJY_usertemp] where id="&cstr(str2(i))
rs.open sql,conn,1,3
next
response.write "<script language=JavaScript>" & chr(13) & "alert('删除成功！');" &"window.location='user_Temporary.asp'" & "</script>"
response.end
Call CloseRs(rs)
End If%>
<FORM name=thisForm action="user_Temporary.asp?dick=del" method=POST>
<div align="center">
  <center>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%" height="130">
<TR>
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">临时会员管理</FONT></TD>
 </TR>
 <TR>
    <TD height="20" bgcolor="#ffffff" colspan="13">&nbsp;&nbsp;&nbsp;&nbsp;当系统开启了会员注册邮件验证后，会员注册在通过邮件验证前的注册资料将保存在此，当通过邮件验证或超过限定时间未进行邮件验证的，将自动删除。网站管理员可在此观察会员注册情况，也可作删除操作。</TD>
 </TR>
  <tr bgcolor="#eeeeee">
   <td width="5%" height="24" align="center" bgcolor="#eeeeee">选择</td>
    <td width="5%" height="24" align="center" bgcolor="#eeeeee">编号</td>
    <td width="10%" height="24" align="center">登录帐号</td>    
    <td width="10%" height="24" align="center">会员姓名</td>
    <td width="15%" height="24" align="center">电子信箱</td> 
    <td width="10%" height="24" align="center">注册时间</td>
	<td width="10%" height="24" align="center">有效时间(天:时:分:秒)</td>
    <td width="5%" height="24" align="center">删除</td>
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
response.write "<table>还没有记录！</table>"
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
    <a href="?dick=del&selectedid=<%=rs("id")%>">删除</a></td>
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
  全选记录
  <input  type="submit" value="删除记录" name="B1">
  共有&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;条记录 共 <font color="#CC5200"><%=Allpage%></font> 页 现在是第 <font color="#CC5200"><%=ThisPage%></font> 页
<%
if ThisPage<2 then     
response.write "<font color=""#808080"">首页</font>&nbsp;"
response.write "<font color=""#808080"">上一页</font>&nbsp;"     
else     
response.write "<a href=?page=1&t1="&trim(request("t1"))&">首页</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&t1="&trim(request("t1"))&">上一页</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">下一页</font>&nbsp;"
response.write "<font color=""#808080"">尾页</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&t1="&trim(request("t1"))&">下一页</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&t1="&trim(request("t1"))&">尾页</a>&nbsp;"     
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
sql="delete from [DNJY_usertemp] where DateDiff("&DN_DatePart_D&",zcdata,"&DN_Now&")>"&regyxq&""'删除3天未邮件认证的临时注册信息
rsdq.open sql,conn,1,3
End If
%>