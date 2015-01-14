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
<%call checkmanage("11")%>
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
<SCRIPT language=JavaScript>
function showoperatealert(id)
{
		if (id==1)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="Admin_Template_Main.asp?dick=hy";
			thisForm.submit();
		}
		}
		if (id==2)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="Admin_Template_Main.asp?dick=del";
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
<%Call OpenConn
Dim id,str2,i
id=request("selectedid")
dick=request("dick")
If dick="hy" Then
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('没有选择记录！');" &"window.location='Admin_Template_Main.asp'" & "</script>"
response.end
end if
str2=split(id,",")
set rs=server.createobject("adodb.recordset")
for i=0 to ubound(str2)
conn.execute "UPDATE DNJY_Template SET M_Reclaim=0 WHERE id="&cstr(str2(i))
next
response.write "<script language=JavaScript>" & chr(13) & "alert('还原成功！');" &"window.location='Admin_Template_Main.asp'" & "</script>"
response.end
Call CloseRs(rs)
End If


If dick="del" Then
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('没有选择记录！');" &"window.location='Admin_Template_Main.asp'" & "</script>"
response.end
end if
str2=split(id,",")
set rs=server.createobject("adodb.recordset")
for i=0 to ubound(str2)
sql="delete from [DNJY_Template] where id="&cstr(str2(i))
rs.open sql,conn,1,3
Next

response.write "<script language=JavaScript>" & chr(13) & "alert('删除成功！');" &"window.location='Admin_Template_Main.asp'" & "</script>"
response.End
Call CloseRs(rs)
End If%>
<FORM name=thisForm action="Admin_Template_Main.asp?dick=del" method=POST>
<div align="center">
  <center>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">

  <TR> 
    <TD height="5" colspan="12" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">模 板 方 案 回 收 站 管 理</FONT></TD>
  </TR>
  <tr bgcolor="#eeeeee">
   <td width="5%" height="24" align="center" bgcolor="#eeeeee">选择</td>
    <td width="5%" height="24" align="center" bgcolor="#eeeeee">编号</td>
    <td width="8%" height="24" align="center">模板方案名称</td>
    <td width="5%" height="24" align="center">还原</td>    
    <td width="5%" height="24" align="center">彻底删除</td>
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
sql = "select * from DNJY_Template where M_Reclaim=1 order by id "&DN_OrderType&""
rs.open sql,conn,1,1
if rs.eof then
response.write "<table>回收站没有模板方案！</table>"
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
    <td  height="20" align="center" bordercolor="#999999" style="border-bottom-style: solid; border-bottom-width: 1"><font color="#FF0000"><%=rs("M_name")%></font></td>    
    <td  height="24" align="center" bordercolor="#999999" style="border-bottom-style: solid; border-bottom-width: 1"><a href="?dick=hy&selectedid=<%=rs("id")%>">还原</a></td>
    <td  height="24" align="center" bordercolor="#999999" style="border-bottom-style: solid; border-bottom-width: 1">
   <INPUT type="button" name="dick" onClick="{if(confirm('确定要删除这个模板方案吗？\n此操作将无法恢复！')){location.href='?selectedid=<%=rs("id")%>&dick=del';}return false;}" value="删除模板方案" ></td>
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
  <INPUT onclick=CheckAll(this.form) type=checkbox value=on name=chkall>选中所有<input onclick=javascript:showoperatealert(1) type="submit" value="全部还原" name="B1">
<input onclick=javascript:showoperatealert(2) type="submit" value="全部删除" name="B2">
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
