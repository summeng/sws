<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("05")%>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
	margin-top: 16px;
}
-->
</style><body leftmargin="0">
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
if (confirm("真的要删除这条记录吗？删除后此记录里的所有内容都将被删除并且无法恢复！"))
window.location = params ;
}
function mm_jumpmenu(targ,selobj,restore){ //v3.0
  eval(targ+".location='"+selobj.options[selobj.selectedindex].value+"'");
  if (restore) selobj.selectedindex=0;
}
//document.all.left_sys.style.display='';
//document.all.left_bm.style.display='none'
-->
</script>
<script language=javascript>
//弹出和管理回复开始
<!--
function left2_menu(meval)
{
  var left_n=eval(meval);
  if (left_n.style.display=="none")
  { eval(meval+".style.display='';"); }
  else
  { eval(meval+".style.display='none';"); }
}
function DoEmpty(params)
{
if (confirm("真的要删除这条记录吗？删除后此记录里的所有内容都将被删除并且无法恢复！"))
window.location = params ;
}
function mm_jumpmenu(targ,selobj,restore){ //v3.0
  eval(targ+".location='"+selobj.options[selobj.selectedindex].value+"'");
  if (restore) selobj.selectedindex=0;
}
//document.all.left_sys.style.display='';
//document.all.left_bm.style.display='none'
-->
</script>
<!---弹出面板开始--->
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
<!---弹出面板结束--->
<!---全选验证删除开始--->
<SCRIPT language=JavaScript>
function showoperatealert(id)
{
		if (id==1)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="info_html_s.asp";
			thisForm.submit();
		}
		}
		if (id==2)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="info_html_sc.asp?type=xinxi";
			thisForm.submit();
		}
		}
		if (id==3)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="info_html_del.asp?urlpage=info_html.asp";
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
<!---全选验证删除结束--->
<div align="center">
  <center>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="99%" height="1" bordercolor="#ADAED6">
  <tr>
    <td width="100%" height="24" colspan="11">
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="25">
  <TR>
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">生 成 html 管 理</FONT></TD>
 </TR>
      <tr>
        <td width="200" height="24">
    <FORM name=sou1 action="?dick=1" method=POST>
    <p align="center">按帐号：<input type="text" name="key1" size="10">
    <input type="submit" value="查找" name="B1">
    </form>
        </td>
        <td width="200" height="24">
    <FORM name=sou2 action="?dick=2" method=POST>
    <p align="center">按姓名：<input type="text" name="key2" size="10">
    <input type="submit" value="查找" name="B1">
    </form>
        </td>
        <td width="360">
    <FORM name=sou3 action="?dick=3" method=POST>
    <p align="center">按关键词：<input type="text" name="key3" size="28"><select name="skey" style="BACKGROUND-COLOR: #FFCCFF; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;">
                            <option value="0">按标题</option>
                            <option value="1">按简介</option>                            
                          </select>
    <input type="submit" value="查找" name="B1">
    </form>
        </td>
      </tr>
	  <tr>
</table>
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="25">
<%
dim k,dick,tj
dim ThisPage,Pagesize,Allrecord,Allpage

   dick=trim(request("dick"))
	city_oneid=Request("city_one")
	city_twoid=Request("city_two")
	city_threeid=Request("city_three")
   If city_oneid="" Then city_oneid=0
   If city_twoid="" Then city_twoid=0
   If city_threeid="" Then city_threeid=0

	IF city_oneid=0 then
	ttcc=""
	elseIF city_threeid>0 then
	ttcc=" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid
	elseIF city_twoid>0 then
	ttcc=" and city_oneid="&city_oneid&" and city_twoid="&city_twoid
	Else
	ttcc=" and city_oneid="&city_oneid
	End If
	
	type_oneid=Request("type_one")
	type_twoid=Request("type_two")
	type_threeid=Request("type_three")
   If type_oneid="" Then type_oneid=0
   If type_twoid="" Then type_twoid=0
   If type_threeid="" Then type_threeid=0

	IF type_oneid=0 then
	tttt=""
	elseIF type_threeid>0 then
	tttt=" and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and type_threeid="&type_threeid
	elseIF type_twoid>0 then
	tttt=" and type_oneid="&type_oneid&" and type_twoid="&type_twoid
	Else
	tttt=" and type_oneid="&type_oneid
	End If%>
	<td height="25" width="20%">按地区类别查询：</td>
        <td width="50%" colspan="4" height="24">
 <form method="POST" name="form" id="form" language="javascript"  action="?t1=<%=trim(request("t1"))%>&t2=<%=trim(request("t2"))%>&t3=<%=trim(request("t3"))%>&dick=<%=dick%>">
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
            <%set rs=conn.execute("select * from DNJY_city where id>0 and twoid=0 order by indexid asc")
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
 <input class="inputb" name="I2" type="submit" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;"  value="查询" border="0" /></font>
</form> </td><%If city_oneid>0 then%><td width="30%" colspan="2" align="left"><%city_one=Conn.Execute("select city from DNJY_city where twoid=0 and id="&city_oneid)(0)
	If city_twoid>0 Then city_two=Conn.Execute("select city from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid)(0)
	If city_threeid<>0 Then city_three=Conn.Execute("select city from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid)(0)
	response.write ""&city_one&"/"&city_two&"/"&city_three&""%></td><%End if%>
      </tr>

<tr>
<form method="POST" name="myform" id="myform" language="javascript"  action="?t1=<%=trim(request("t1"))%>&t2=<%=trim(request("t2"))%>&t3=<%=trim(request("t3"))%>&dick=<%=dick%>">
<td height="25" width="20%">按信息类别查询：</td>
<td height="25" width="50%"><%
set rs=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid=0 order by indexid")
%>
          <SCRIPT language = "JavaScript">
var onecount2;
onecount2=0;
subcat2 = new Array();
        <%
		dim count2:count2 = 0
        do while not rs.eof 
        %>
subcat2[<%=count2%>] = new Array("<%=rs("name")%>","<%=rs("id")%>","<%=rs("twoid")%>");
        <%
        count2 = count2 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount2=<%=count2%>;
          </SCRIPT>
          <%
set rs=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid>0 order by indexid")
%>
          <SCRIPT language = "JavaScript">
var onecount3;
onecount3=0;
subcat3 = new Array();
        <%dim count3:count3 = 0
        do while not rs.eof%>
subcat3[<%=count3%>] = new Array("<%=rs("name")%>","<%=rs("id")%>","<%=rs("twoid")%>","<%=rs("threeid")%>");
        <%
        count3 = count3 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount3=<%=count3%>;

function changelocation2(locationid)
    {
    document.myform.type_two.length = 0; 
    document.myform.type_two.options[0] = new Option('二级分类','');
	document.myform.type_three.length = 0; 
    document.myform.type_three.options[0] = new Option('三级分类','');
    var locationid=locationid;
	var i;
    for (i=0;i < onecount2; i++)
        {
            if (subcat2[i][1] == locationid)
            { 
                document.myform.type_two.options[document.myform.type_two.length] = new Option(subcat2[i][0], subcat2[i][2]);
            }        
        }
       
    }    
     function changelocation3(locationid,locationid1)
    {
    document.myform.type_three.length = 0; 
    document.myform.type_three.options[0] = new Option('三级分类','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount3; i++)
        {
            if (subcat3[i][2] == locationid)
            { 
			if (subcat3[i][1] == locationid1)
			{
                document.myform.type_three.options[document.myform.type_three.length] = new Option(subcat3[i][0], subcat3[i][3]);
            }        
        }
       } 
    }            
	      </SCRIPT>
          <SELECT name="type_one" size="1" id="select8" onChange="changelocation2(document.myform.type_one.options[document.myform.type_one.selectedIndex].value)">
            <OPTION value="" selected>一级分类</OPTION>
            <%set rs=conn.execute("select * from DNJY_type where id>0 and twoid=0 order by indexid asc")
if rs.eof or rs.bof then
response.write "<option value=''>没有分类</option>"
else
do until rs.eof
response.write "<option value='"&rs("id")&"'>"&rs("name")&"</option>"
    rs.movenext
    loop
end if
Call CloseRs(rs)%>
          </SELECT>
          <SELECT name="type_two" id="select9" onChange="changelocation3(document.myform.type_two.options[document.myform.type_two.selectedIndex].value,document.myform.type_one.options[document.myform.type_one.selectedIndex].value)">
            <OPTION value="" selected>二级分类</OPTION>
          </SELECT>
          <SELECT name="type_three" id="select10">
            <OPTION value="" selected>三级分类</OPTION>
          </SELECT>  <input class="inputb" name="t2" type="submit" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;"  value="查询" border="0" /></font>
</form></td><%If type_oneid>0 then%><td width="30%" colspan="2" align="left"><%type_one=Conn.Execute("select name from DNJY_type where twoid=0 and id="&type_oneid)(0)
	If type_twoid>0 Then type_two=Conn.Execute("select name from DNJY_type where id="&type_oneid&" and threeid=0 and twoid="&type_twoid)(0)
	If type_threeid<>0 Then type_three=Conn.Execute("select name from DNJY_type where id="&type_oneid&" and twoid="&type_twoid&" and threeid="&type_threeid)(0)
	response.write ""&type_one&"/"&type_two&"/"&type_three&""%></td><%End if%>
                        </tr>
      <tr>
        <td width="100%" colspan="11" height="24">
        <p align="center"><b>显示：</b><a href="?dick=4&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>">未(<font color="#FF0000">ⅹ</font>)验证信息</a>---<a href="?dick=5&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>">已(√)验证信息</a>---<a href="?dick=6&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>">已生成html信息</a>---<a href="?dick=7&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>">未生成html信息</a><%If HtmlHome=1 then%><br><a href="info_makeHtml.asp?mudi=home"><font color="#0000FF">生成首页静态页</font></a><font color="#FF0000">(在“基本参数”设置中当选择按分站方式显示时请勿生成，否则不能按地区显示信息！)</font><%End if%>&nbsp;&nbsp;&nbsp;&nbsp;（注：已设置浏览权限的信息必须动态页面显示，不能生成静态页，不在此显示）</td>
      </tr>
    </table>
    </td>
  </tr>
  <FORM name=thisForm action="" method=POST>
<%
function HTMLEncode(fString)
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
end Function


set rs = Server.CreateObject("ADODB.RecordSet")
Select Case dick
Case "1"
sql="select * from [DNJY_info] where Readinfo=0 and username='"&trim(request("key1"))&"'"&ttcc&""&tttt&" order by id "&DN_OrderType&"" 
Case "2"
sql="select * from [DNJY_info] where Readinfo=0 and name like '%"&trim(request("key2"))&"%'"&ttcc&""&tttt&" order by id "&DN_OrderType&"" 
Case "3"
If request("skey")=0 then
sql="select * from [DNJY_info] where Readinfo=0 and (biaoti like '%"&trim(request("key3"))&"%')"&ttcc&""&tttt&" order by id "&DN_OrderType&""
Else
sql="select * from [DNJY_info] where Readinfo=0 and (memo like '%"&trim(request("key3"))&"%')"&ttcc&""&tttt&" order by id "&DN_OrderType&""
End if 
Case "4"
sql="select * from [DNJY_info] where Readinfo=0 and yz<>1"&ttcc&""&tttt&" order by id "&DN_OrderType&"" 
Case "5"
sql="select * from [DNJY_info] where Readinfo=0 and yz=1"&ttcc&""&tttt&" order by id "&DN_OrderType&""
Case "6"
sql="select * from [DNJY_info] where Readinfo=0 and fsohtm=1"&ttcc&""&tttt&" order by id "&DN_OrderType&""
Case "7"
sql="select * from [DNJY_info] where Readinfo=0 and fsohtm=0"&ttcc&""&tttt&" order by id "&DN_OrderType&""
Case Else
sql="select * from [DNJY_info] where Readinfo=0 and fbts>0"&ttcc&""&tttt&" order by id "&DN_OrderType&"" 
End Select



rs.open sql,conn,1,1
if rs.eof then
response.write "<li>没有记录"
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
  <tr bgcolor="#BDBEDE">
    <td width="5%" height="24" align="center" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">选择</td>
   <td width="5%" height="24" align="center" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">ID</td>
    <td width="25%" height="28" bordercolor="#666666" bgcolor="#BDBEDE" style="border-bottom-style: solid; border-bottom-width: 1">标题</td>
    <td width="13%" height="24" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">
    <p align="center">注册ID/姓名</td>
    
    <td width="5%" height="24" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">
    <p align="center">审核状态</td>    
	<td width="5%" height="24" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">
    <p align="center">已生成</td>  
    <td width="20%" height="24" bordercolor="#666666" bgcolor="#BDBEDE" style="border-bottom-style: solid; border-bottom-width: 1"><div align="center">生成HTML操作</div></td>
  </tr>
<%dim uid
k=0
do while not rs.eof
uid=rs("id")

	    dim rs1,sql1
		set rs1=server.createobject("adodb.recordset") 
		sql1="select * from DNJY_reply where xxid="&rs("id")&" order by id "&DN_OrderType&""
		rs1.open sql1,conn,1,1
		
		dim rs2,sql2
		set rs2=server.createobject("adodb.recordset") 
		sql2="select * from DNJY_Bad_info where infoid="&rs("id")&" order by id "&DN_OrderType&""
		rs2.open sql2,conn,1,1%>
  <tr>
    <td width="5%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1">
    <INPUT type=checkbox value="<%=uid%>" name=selectedid></td>
    <td width="5%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1">
    <%= rs("id") %></td>
    <td width="20%" height="24" bgcolor="#FAFAFA" style="border-top-style: solid; border-top-width: 1">
   <% if rs("tupian")<>"0" then
response.write "<img src=""../images/num/pic.gif"">" 
end if%>

<%	belongtype=""&rs("type_one")&""
	IF rs("type_two")<>"" and not isnull(rs("type_two")) Then belongtype=belongtype&" / "&rs("type_two")&""
	IF rs("type_three")<>"" and not isnull(rs("type_three")) Then belongtype=belongtype&" / "&rs("type_three")&""
	%><A href="admin_info.asp?id=<%=uid%>" target="_blank" title="信息类别：<%=belongtype%>"><%=rs("biaoti")%></a></td>
    <td width="13%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1"><a target="_blank" href=../preview.asp?username=<%= rs("username") %>><%=rs("username")%><font color="#FF0000">/</font><%=rs("name")%></td>
    
    <td width="5%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1"><%if rs("yz")=1 then%>√<%else%><b><font color="#FF0000">ⅹ</font></b><%end if%></td>
    
    <td width="5%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1"><%if rs("fsohtm")=1 then%><span title=文件位置：/html/<%=rs("id")%>.html style="cursor:help">√</span><%else%><b><font color="#FF0000">ⅹ</font></b><%end if%></td>    
    <td width="25%" height="24" bgcolor="#FAFAFA" style="border-top-style: solid; border-top-width: 1"><%if rs("yz")=1 then%> <A href="info_html_s.asp?selectedid=<%=rs("id")%>" title='生成本信息的HTML文件'>生成</a> <%if rs("fsohtm")=1 then%><A href="../html/<%=rs("id")%>.html" target="_blank" title='查看本信息的HTML页面'>查看</a> <A href="info_html_del.asp?selectedid=<%=rs("id")%>&qxyz=no&page=<%=(ThisPage)%>&t1=<%=trim(request("t1"))%>&t2=<%=trim(request("t2"))%>&t3=<%=trim(request("t3"))%>&dick=<%=dick%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>&urlpage=info_html.asp" title='删除本信息的HTML文件'>删除htm文件</a><%else%><font color="#FF0000">尚未生成</font><%end if%><%end if%> <a href="info_del.asp?selectedid=<%=uid%>&urlpage=info_html.asp">删除该信息</a> <A href="info_edit.asp?id=<%=uid%>">编辑该信息</a></td>
  </tr>

<%
k=k+1
rs.movenext
if k>=Pagesize then exit do
loop
Call CloseRs(rs)

%>
  <tr>
    <td width="100%" height="1" colspan="11" style="border-top-style: solid; border-top-width: 1">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="1">
<tr> 
<td width="100%" height="25" align="center" bgcolor="#eeeeee"><div align="center">
      <input onClick=CheckAll(this.form) type=checkbox value=on name=chkall>
      选中所有 <input onclick=javascript:showoperatealert(1) type="submit" value="生成选定信息(已审核)的htm文件" name="B1">	  	  
       <input onclick=javascript:showoperatealert(3) type="submit" value="删除选定信息的htm文件" name="B3" style="color: #FF0000">



<br>共有&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;条记录 共 <font color="#CC5200"><%=Allpage%></font> 页 现在是第 <font color="#CC5200"><%=ThisPage%></font> 页
    <%
if ThisPage<2 then     
response.write "<font color=""#808080"">首页</font>&nbsp;"
response.write "<font color=""#808080"">上一页</font>&nbsp;"     
else     
response.write "<a href=?page=1&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&type_one="&type_oneid&"&type_two="&type_twoid&"&type_three="&type_threeid&">首页</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&type_one="&type_oneid&"&type_two="&type_twoid&"&type_three="&type_threeid&">上一页</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">下一页</font>&nbsp;"
response.write "<font color=""#808080"">尾页</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&type_one="&type_oneid&"&type_two="&type_twoid&"&type_three="&type_threeid&">下一页</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&type_one="&type_oneid&"&type_two="&type_twoid&"&type_three="&type_threeid&">尾页</a>&nbsp;"     
end if
%>
</div></td>
</tr>
</table>
</form>
<tr>
<td width="100%" height="1" colspan="11" style="border-top-style: solid; border-top-width: 1">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="1">
<FONT color="#FF0000">注：根据数据量多少点击开始生成后要等待一定时间。建议每次生成100条记录内，否则可能造成资源不足而死机</font>
<%Dim oneid,twoid,threeid
oneid=strint(Request("oneid"))
twoid=strint(Request("twoid"))
threeid=strint(Request("threeid"))
    Response.Write "        <tr>"
    Response.Write "          <td height='40'>"
    Response.Write "            <form name='form1' method='post' action='info_html_s.asp'>"
    Response.Write "            生成最新 <input name='TopNew' id='TopNew' value='10' size=8 maxlength='10'> 条信息（建议10条以下）&nbsp;"
    Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='1'>"  
    Response.Write "            <input name='submit2' type='submit' id='submit2' value='开始生成>>'>"
    Response.Write "            </form>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td height='40'>"
    Response.Write "            <form name='form2' method='post' action='info_html_sc.asp?type=xinxi'>"
    Response.Write "            生成ID号为 <input name='BeginID' type='text' id='BeginID' value='1' size=8 maxlength='10'> 到 <input name='EndID' type='text' id='EndID' value='30' size=8 maxlength='10'> 的信息（建议30条以下）"
     Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='2'>"  
    Response.Write "            <input name='submit2' type='submit' id='submit2' value='开始生成>>'>"
    Response.Write "            </form>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td height='40'>"
    Response.Write "            <form name='form3' method='post' action='info_html_s.asp'>"
    Response.Write "            生成指定ID的信息（多个ID可用半角英文“,”号隔开）："
    Response.Write "            <input name='selectedid' type='text' id='selectedid' value='1,3,5' size='20'>" 
    Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='3'>"  
    Response.Write "            <input name='submit2' type='submit' id='submit2' value='开始生成>>'>"
    Response.Write "            </form>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td height='40'>"
    Response.Write "            <form name='form4' method='post' action='info_html_sc.asp?type=xinxi&oneid="&oneid&"&twoid="&twoid&"&threeid="&threeid&"'>"
    Response.Write "            生成栏目名称为："%>
			  <SELECT name="name" onChange="location=this.value">
			     <OPTION value="?oneid=<%=Rs(0)%>" selected>选择栏目</OPTION>
				  <%Set Rs=Conn.Execute("Select id,twoid,threeid,name from DNJY_type order by id,twoid,threeid")
				  Do While Not Rs.Eof
					  If Rs(0)<>0 and Rs(1)=0 and Rs(2)=0 Then%>
					  <OPTION style="background-color:#F8F4F5 ;color: #FF0000" value="?oneid=<%=Rs(0)%>" <%if Rs(0)=oneid and 0=twoid Then%>selected<%End if%>><%=Rs(3)%></OPTION>
					  <%end if
					  If Rs(0)<>0 and Rs(1)<>0 and Rs(2)=0 Then%>
					  <OPTION style="background-color:#F8F4F5 ;color: #0000FF" value="?oneid=<%=Rs(0)%>&twoid=<%=Rs(1)%>" <%if Rs(0)=oneid and Rs(1)=twoid and 0=threeid Then%>selected<%End if%>>├<%=Rs(3)%></OPTION>
					  <%end if
					  If Rs(0)<>0 and Rs(1)<>0 and Rs(2)<>0 Then%>
					  <OPTION value="?oneid=<%=Rs(0)%>&twoid=<%=Rs(1)%>&threeid=<%=Rs(2)%>" <%if Rs(0)=oneid and Rs(1)=twoid and Rs(2)=threeid Then%>selected<%End if%>>├├<%=Rs(3)%></OPTION>
					  <%End if
					  Rs.MoveNext
				  Loop
				  %>
              </SELECT>的信息
 <%Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='4'>"  
    Response.Write "            <input name='submit2' type='submit' id='submit2' value='开始生成>>'>"
    Response.Write "            </form>"
    Response.Write "          </td>"
    Response.Write "        </tr>"

    Response.Write "        <tr>"
    Response.Write "          <td height='40'>"
    Response.Write "            <form name='form7' method='post' action='info_html_sc.asp?type=username'>"
    Response.Write "            生成会员名称为："
    Response.Write "            <input name='name' type='text' id='name' value='' size='20'> 的信息" 
    Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='7'>"  
    Response.Write "            <input name='submit2' type='submit' id='submit2' value='开始生成>>'>"
    Response.Write "            </form>"
    Response.Write "          </td>"
    Response.Write "        </tr>"

    Response.Write "        <tr>"
    Response.Write "          <td height='40'>"
    Response.Write "            <form name='form5' method='post' action='info_html_sc.asp?type=xinxi'>"
    Response.Write "            <font color='red'>生成所有未生成的信息（推荐）</font>"
    Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='5'>"  
    Response.Write "            <input name='submit2' type='submit' id='submit2' value='开始生成>>'>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='Admin_UpdateCreatedStatus.asp' title='当您用FTP工具删除了已经生成的HTML页面，或者换了服务器后，可以用此功能更新数据库记录的生成状态，以能正常使用“生成所有未生成的信息”功能。'>【更新信息生成状态】</a>"
    Response.Write "            </form>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td height='40'>"
    Response.Write "            <form name='form6' method='post' action='info_html_sc.asp?type=xinxi'>"
    Response.Write "            生成所有审核通过的信息"
    Response.Write "            <input name='CreateType' type='hidden' id='CreateType' value='6'>"  
    Response.Write "            <input name='submit2' type='submit' id='submit2' value='开始生成>>'>（全部连续生成非常耗资源，慎重使用！）"
    Response.Write "            </form>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
	Call CloseDB(conn)%>
</td>
</tr>
</table>
    </td>
  </tr>

</table></center>
</div>
