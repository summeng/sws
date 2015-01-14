<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("07")%>
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
.style3 {font-size: 12px}
-->
</style>
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
</script>
<BODY>
<div align="center">
  <center>

 <%Dim key,known,str2
	city_oneid=Request("city_one")
	city_twoid=Request("city_two")
	city_threeid=Request("city_three")
   If city_oneid="" Then city_oneid=0
   If city_twoid="" Then city_twoid=0
   If city_threeid="" Then city_threeid=0%>
<table border="1" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%" cellpadding="0">
  <tr bgcolor="#BDBEDE">
    <td width="100%" height="28" colspan="8"><p class="style3"><b><span class="style1">&nbsp;</span><span class="style2">留言管理</span></b></td>
  </tr>
  <tr>
    <td width="100%" height="23" colspan="3"> <b> <a href="?dick=3&page=<%=Page%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>">
    <font size="2"><span class="style3"><b><span class="style1">&nbsp;</span></b></span></font></a></b><a href="?dick=3&page=<%=Page%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>"><font size="2">列出所有留言</font></a><font size="2"> <a href="?dick=1&page=<%=Page%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>">列出投诉留言</a>
    <a href="?dick=2&page=<%=Page%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>">列出普通留言</a> </font> <a href="?dick=4&page=<%=Page%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%> ">
    <font size="2">没有回复的留言</font></a></td>  <td colspan="5"> 
                                  
 
<%
key=Request("key")
id=trim(request("id"))
If key="del" then
if id="" then
response.write "<li>没有选择记录！"
response.end
end If
str2=split(id,",")
for i=0 to ubound(str2)'循环开始
set rs=server.createobject("adodb.recordset")
sql="delete from [DNJY_gbook] where id="&cstr(str2(i))
rs.open sql,conn,1,3
set rs=Nothing
Next'循环结束
response.Write "<script language=javascript>alert('删除成功!');location.replace('?page="&page&"&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"');</script>"
end if

If key="kg" then
known=trim(request("known"))
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_gbook] where id="&trim(id)
rs.open sql,conn,1,3
If known="no" then
rs("known")=1
Else
rs("known")=0
End if
rs.update
Call CloseRs(rs)
end If

%>
<form method="POST" name="form" id="myform" language="javascript"  action="?">
<%
set rs=server.createobject("adodb.recordset")
set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")%>
<SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
	dim count5:count5 = 0
        do while not rs.eof 
        %>
subcat[<%=count5%>] = new Array("<%=rs("city")%>","<%=rs("id")%>","<%=rs("twoid")%>");
        <%
        count5 = count5 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount=<%=count5%>;
          </SCRIPT>
          <%
set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid>0")
%>
          <SCRIPT language = "JavaScript">
var onecount4;
onecount4=0;
subcat4 = new Array();
        <%
		dim count4:count4 = 0
        do while not rs.eof 
        %>
subcat4[<%=count4%>] = new Array("<%=rs("city")%>","<%=rs("id")%>","<%=rs("twoid")%>","<%=rs("threeid")%>");
        <%
        count4 = count4 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount4=<%=count4%>;

function changelocation(locationid)
    {
    document.form.city_two.length = 0; 
	document.form.city_two.options[0] = new Option('选择城市','');
	document.form.city_three.length = 0; 
	document.form.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.form.city_two.options[document.form.city_two.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.form.city_three.length = 0; 
    document.form.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.form.city_three.options[document.form.city_three.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
          <SELECT name="city_one" size="1" id="select" onChange="changelocation(document.form.city_one.options[document.form.city_one.selectedIndex].value)">
            <OPTION value="" selected>选择地区</OPTION>
            <%set rs=conn.execute("select * from DNJY_city where id>0 and twoid=0  order by indexid asc")
if rs.eof or rs.bof then
response.write "<option value=''>没有分类</option>"
else
do until rs.eof
response.write "<option value='"&rs("id")&"'>"&rs("city")&"</option>"
    rs.movenext
    loop
end if
Call CloseRs(rs)
%>
          </SELECT>
          <SELECT name="city_two" id="select6" onChange="changelocation4(document.form.city_two.options[document.form.city_two.selectedIndex].value,document.form.city_one.options[document.form.city_one.selectedIndex].value)">
            <OPTION value="" selected>选择城市</OPTION>
          </SELECT>
          <SELECT name="city_three" id="select7">
            <OPTION value="" selected>选择城市</OPTION>
          </SELECT>
 <input class="inputb" name="I2" type="submit" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;"  value="查询" border="0" /></font></td>
  </tr>
</form>
<%
dim k,dick
Dim Pagesize,Allrecord,Allpage,ThisPage,page
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end if
dick=trim(request("dick"))
 If dick="" Then dick=0
set rs=server.createobject("adodb.recordset")
	IF city_oneid=0 then
	ttccbook=""
	elseIF city_threeid>0 then
	ttccbook=" city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid
	elseIF city_twoid>0 then
	ttccbook=" city_oneid="&city_oneid&" and city_twoid="&city_twoid
	Else
	ttccbook=" city_oneid="&city_oneid
	End If
	If shows5=0 Then ttccbook=""
If city_oneid>0 Then
sql = "select * from DNJY_gbook where "&ttccbook&" order by id "&DN_OrderType&""
else
Select Case dick
Case "1"
sql = "select * from DNJY_gbook where lx=1 order by id "&DN_OrderType&""
Case "2"
sql = "select * from DNJY_gbook where lx=0 order by id "&DN_OrderType&""
Case "4"
sql = "select * from DNJY_gbook where hf=0 order by id "&DN_OrderType&""
Case Else
sql = "select * from DNJY_gbook order by id "&DN_OrderType&""
End Select
end if
rs.open sql,conn,1,1
if rs.eof Then
response.write "<tr><td width=""40%"" colspan=""4"">暂无"
IF city_oneid>0 Then response.write conn.Execute("Select city From DNJY_city Where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid&"")(0)
response.write "留言</td></tr>"
Else

rs.Pagesize=15
Pagesize=rs.Pagesize
Allrecord=rs.Recordcount
Allpage=rs.Pagecount
if ThisPage<1 then                           
ThisPage=1
end if
On Error Resume Next
rs.move (ThisPage-1)*Pagesize
%>
  <tr bgcolor="#eeeeee">
   <td width="6%" height="23"><div align="center">选择</div></td>
    <td width="6%" height="23"><div align="center">编号</div></td>
    <td width="50%" height="23"><div align="center">内容</td>
    <td width="8%" height="23"><div align="center">地区</td>
	<td width="8%" height="23"><div align="center">是否公开</td>
    <td width="6%" height="23" align="center">　</td>
    <td width="6%" height="23" align="center">　</td>
    <td width="15%" height="23" align="center">是否回复</td>
  </tr>
<form name="form1" method="post" action="?key=del&&page=<%=ThisPage%>&city_one=<%=city_oneid%>&city_two=<%city_twoid%>&city_three=<%city_threeid%>">
  <%k=1
  do while not rs.eof%>
  <tr bgcolor="#FFFFFF">
    <td width="4%" height="23" bgcolor="#FAFAFA">
    <p align="center"><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>"></td>
    <td width="4%" height="23" bgcolor="#FAFAFA">
    <p align="center"><%=k%></td>
    <td width="50%" height="23"><%=rs("gbook1")%><br>〈<font color="#999999"><%=rs("username")%></font>&nbsp;<font color="#999999"><%=rs("fbsj")%></font>〉</td>
     <TD width="8%" align="center"><A title=<%=rs("city_one")%>/<%=rs("city_two")%>/<%=rs("city_three")%> href="#"><%=rs("city_one")%></A></TD>
	 <td width="8%" height="23" align="center" bgcolor="#FAFAFA"><%if rs("known")=0 then%><a href="?known=no&id=<%=rs("id")%>&key=kg&page=<%=ThisPage%>&city_one=<%=city_oneid%>&city_two=<%city_twoid%>&city_three=<%city_threeid%>">√</a><%else%><b><a href="?known=yes&id=<%=rs("id")%>&key=kg&page=<%=ThisPage%>&city_one=<%=city_oneid%>&city_two=<%city_twoid%>&city_three=<%city_threeid%>"><font color="#FF0000">ⅹ</font></a></b><%end if%></td>
    <td width="6%" height="23" align="center" bgcolor="#FAFAFA"><a href="#" ONCLICK="window.open('bookedit.asp?id=<%=rs("id")%>&chk=2','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=250,height=300,left=300,top=100')">
    <font color="#FF0000">修改</font></a></td>
    <td width="6%" height="23" align="center"><a href="javascript:DoEmpty('?key=del&id=<%=rs("id")%>&page=<%=ThisPage%>&city_one=<%=city_oneid%>&city_two=<%city_twoid%>&city_three=<%city_threeid%>')"><font color="#FF0000">删除</font></a></td>
    <td width="11%" height="23" align="center" bgcolor="#FAFAFA"><%if rs("hf")=1 then%><span id="followImg<%=k%>" style="CURSOR: hand" onclick="loadThreadFollow(<%=k%>,5)">查看回复</span><%else%><a href="#" ONCLICK="window.open('bookreply.asp?id=<%=rs("id")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=250,height=300,left=300,top=100')"><font color="#FF0000">回复留言</font></a><%end if%></td>
  </tr>
  <tr style="display:none" id="follow<%=k%>">
    <td width="100%" height="23" colspan="7">
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
      <tr>
        <td bgcolor="#CCCCFF"><font color="#FF6600">&nbsp;<%=rs("gbook2")%></font> <%=rs("hfsj")%>&nbsp;[<a href="#" ONCLICK="window.open('bookreply.asp?id=<%=rs("id")%>&chk=2','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=250,height=300,left=300,top=100')">修改</a>]<span class="style3"><b><span class="style2"></span></b></span></td>
        </tr>
    </table>
    </td>
  </tr>
   <%
   rs.movenext
if k>=Pagesize then exit do
   k=k+1
   Loop
   end if
   rs.close
   set rs=nothing
   Call CloseDB(conn)
   %>
</table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
<tr>
<td height="20" colspan="15" align="center" bgcolor="#F3F3F3">	
  <p align="center">
    <input type="checkbox" name="chkall" onClick="CheckAll(this.form)">
    全部选中 
    <input type="submit" name="Submit" value="删除所选留言" onClick="{if(confirm('该操作不可恢复！\n\n确实删除选定的留言？')){return true;}return false;}">
  </p>
 </td> 
</tr>
</form>
</table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
<tr> 
<td width="595" height="25" align="center">
  共有&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;条记录 共 <font color="#CC5200"><%=Allpage%></font> 页 现在是第 <font color="#CC5200"><%=ThisPage%></font> 页
<%
if ThisPage<2 then     
response.write "<font color=""#808080"">首页</font>&nbsp;"
response.write "<font color=""#808080"">上一页</font>&nbsp;"     
else     
response.write "<a href=?page=1&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&">首页</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&">上一页</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">下一页</font>&nbsp;"
response.write "<font color=""#808080"">尾页</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&">下一页</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&">尾页</a>&nbsp;"     
end if
%></td>
</tr>
      </table>   
</center>
</div>
