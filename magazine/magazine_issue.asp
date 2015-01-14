<!--#include file="../conn.asp"-->
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
<%call checkmanage("8")%>

<SCRIPT language=javascript>
<!--
function CheckForm()
{
if(document.formadd.issue.value.length<1)
	{
	    alert("请填写期号!");
        document.formadd.issue.focus();
	    return false;
	}
if(document.formadd.Releasetime.value.length<3)
	{
	    alert("发布时间没有填写!");
        document.formadd.Releasetime.focus();
        document.formadd.Releasetime.select();
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
<%Call OpenConn
Dim rs_b,sql_b,rs_c,sql_c,city_oneid,city_twoid,city_threeid,city_one,city_two,city_three,Releasetime,issue,state
city_oneid=strint(request("city_oneid"))
city_twoid=strint(request("city_twoid"))
city_threeid=strint(request("city_threeid"))
state=request("state")
set rs=server.createobject("adodb.recordset")
if city_oneid>0 then
rs.open "select city from DNJY_city where twoid=0 and threeid=0 and id="&city_oneid,conn,1,1
city_one=rs("city")
rs.close
end if
if city_twoid>0 then
rs.open "select city from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid,conn,1,1
city_two=rs("city")
rs.close
end if
if city_twoid>0 and city_threeid>0 then
rs.open "select city from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid,conn,1,1
city_three=rs("city")
rs.close
end if

if request("action")="add" then
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_zz_issue"
rs.open sql,conn,1,3
rs.addnew
rs("city_oneid")=city_oneid
rs("city_one")=city_one
rs("city_twoid")=city_twoid
rs("city_two")=city_two
rs("city_threeid")=city_threeid
rs("city_three")=city_three
rs("issue")=trim(request("issue"))
rs("Releasetime")=trim(request("Releasetime"))
rs("addtime")=now()
rs("state")=true
rs.update
Call CloseRs(rs)
Conn.Execute("Update DNJY_city Set magazine="&DN_True&" where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid&"")
Call Alert ("添加成功！","magazine_issue.asp")
end if

if request("action")="editchk" then
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_zz_issue where id="&request("id")&""
rs.open sql,conn,1,3
rs("city_oneid")=city_oneid
rs("city_one")=city_one
rs("city_twoid")=city_twoid
rs("city_two")=city_two
rs("city_threeid")=city_threeid
rs("city_three")=city_three
rs("issue")=trim(request("issue"))
rs("Releasetime")=trim(request("Releasetime"))
rs("addtime")=now()
rs("state")=state
rs.update
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_zz_edition where issueid="&request("id")&""
rs.open sql,conn,1,3
do while not rs.eof
rs("city_oneid")=city_oneid
rs("city_one")=city_one
rs("city_twoid")=city_twoid
rs("city_two")=city_two
rs("city_threeid")=city_threeid
rs("city_three")=city_three
rs("addtime")=now()
rs.update
rs("state")=state
rs.movenext
loop
Call CloseRs(rs)
Conn.Execute("Update DNJY_city Set magazine="&DN_True&" where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid&"")
Call Alert ("修改成功！","magazine_issue.asp")
end if

if request("action")="del" then
set rs_b=server.createobject("adodb.recordset")
sql_b = "delete from DNJY_zz_issue where id="&request("id")&""
rs_b.open sql_b,conn,1,3
Conn.Execute("Update DNJY_city Set magazine="&DN_False&",state="&DN_False&" where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid&"")
Call Alert ("删除成功！","magazine_issue.asp")
end if

if request("action")="alldel" Then
'===============删除相关图片===========================
Dim rsdel,ImageName
set rsdel = server.CreateObject ("Adodb.recordset")
sql="select thumbnail,largerpic from DNJY_zz_edition where city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&""
rsdel.open sql,conn,1,1
do while not rsdel.eof
if Not IsNull(rsdel("thumbnail")) Then
   if Not rsdel.eof Then ImageName="/"&strInstallDir&rsdel("thumbnail")
   call DoDelFile(ImageName)  
end If
if Not IsNull(rsdel("largerpic")) Then
   if Not rsdel.eof Then ImageName="/"&strInstallDir&rsdel("largerpic")
   call DoDelFile(ImageName)  
end If
rsdel.movenext
loop
'===============================================
set rs_b=server.createobject("adodb.recordset")
sql_b = "delete from DNJY_zz_edition where city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&""
rs_b.open sql_b,conn,1,3'删除相关版面
set rs_c=server.createobject("adodb.recordset")
sql_c = "delete from DNJY_zz_issue where city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&""
rs_c.open sql_c,conn,1,3'删除相关期刊号
Conn.Execute("Update DNJY_city Set magazine="&DN_False&",state="&DN_False&" where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid&"")'去除杂志标记
Call Alert ("删除成功！","magazine_issue.asp")
end If

if request("action")="stateno" Then
Conn.Execute("Update DNJY_city Set state="&DN_False&" where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid&"")
Conn.Execute("Update DNJY_zz_edition Set state="&DN_False&" where city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&"")
Conn.Execute("Update DNJY_zz_issue Set state="&DN_False&" where city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&"")
ElseIf request("action")="stateok" Then
Conn.Execute("Update DNJY_city Set state="&DN_True&" where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid&"")
End if
%>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
<body topmargin="0" leftmargin="0"><center>
<%if request("action")="edit" then
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_zz_issue where id="&request("id")&""
rs.open sql,conn,1,3
city_oneid=rs("city_oneid")
city_one=rs("city_one")
city_twoid=rs("city_twoid")
city_two=rs("city_two")
city_threeid=rs("city_threeid")
city_three=rs("city_three")
issue=rs("issue")
Releasetime=rs("Releasetime")
Call CloseRs(rs)
%>
<form name="formadd" method="POST" action="?action=editchk&id=<%=request("id")%>"><div align="center">
<table width="98%" border="0" class="table" cellspacing="1" cellpadding="0">
<tr> 
<td height="25" bgcolor="#799AE1" class="tdbar">
&nbsp;分站版修改</td>
</tr>
<tr> 
<td align="center" class="td">
<p>（<a href="../admin/city_Level1.asp">管理地区</a>） &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
所属地区：<%
Dim rsi
set rsi=server.createobject("adodb.recordset")
set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")
%>
<SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
		//dim count:
		count = 0
        do while not rsi.eof 
        %>
subcat[<%=count%>] = new Array("<%=rsi("city")%>","<%=rsi("id")%>","<%=rsi("twoid")%>");
        <%
        count = count + 1
        rsi.movenext
        loop
        rsi.close
		set rsi=nothing
        %>
onecount=<%=count%>;
                             </SCRIPT>
                                   <%
set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid>0")
%>
                                   <SCRIPT language = "JavaScript">
var onecount4;
onecount4=0;
subcat4 = new Array();
        <%
		count4 = 0
        do while not rsi.eof 
        %>
subcat4[<%=count4%>] = new Array("<%=rsi("city")%>","<%=rsi("id")%>","<%=rsi("twoid")%>","<%=rsi("threeid")%>");
        <%
        count4 = count4 + 1
        rsi.movenext
        loop
        rsi.close
		set rsi = nothing
        %>
onecount4=<%=count4%>;

function changelocation(locationid)
    {
    document.formadd.city_twoid.length = 0; 
	document.formadd.city_twoid.options[0] = new Option('选择城市','');
	document.formadd.city_threeid.length = 0; 
	document.formadd.city_threeid.options[0] = new Option('选择城市','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.formadd.city_twoid.options[document.formadd.city_twoid.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.formadd.city_threeid.length = 0; 
    document.formadd.city_threeid.options[0] = new Option('选择城市','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.formadd.city_threeid.options[document.formadd.city_threeid.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
                             </SCRIPT>
                                   <SELECT name="city_oneid" size="1" id="select2" onChange="changelocation(document.formadd.city_oneid.options[document.formadd.city_oneid.selectedIndex].value)">
                                     <OPTION value="" selected>选择城市</OPTION>
                                     <%set rsi=conn.execute("select * from DNJY_city where id>0 and twoid=0 order by indexid asc")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
                                     <OPTION value="<%=rsi("id")%>" <%if rsi("id")=city_oneid then%>selected<%end if%>><%=rsi("city")%></OPTION>
                                     <% rsi.movenext
    loop
end if
rsi.close
set rsi = nothing
%>
                                   </SELECT>
                                   <SELECT name="city_twoid" id="city_twoid" onChange="changelocation4(document.formadd.city_twoid.options[document.formadd.city_twoid.selectedIndex].value,document.formadd.city_oneid.options[document.formadd.city_oneid.selectedIndex].value)">
                                     <OPTION value="" selected>选择城市</OPTION>
                                     <%set rsi=conn.execute("select * from DNJY_city where id="&city_oneid&" and twoid<>0 and threeid=0")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
                                     <OPTION value="<%=rsi("twoid")%>" <%if rsi("twoid")=city_twoid then%>selected<%end if%>><%=rsi("city")%></OPTION>
                                     <% rsi.movenext
    loop
	end if
rsi.close
set rsi = nothing%>
                                   </SELECT>
                                   <SELECT name="city_threeid" id="city_threeid">
                                     <OPTION value="" selected>选择城市</OPTION>
                                     <%set rsi=conn.execute("select * from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid<>0")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
                                     <OPTION value="<%=rsi("threeid")%>" <%if rsi("threeid")=city_threeid then%>selected<%end if%>><%=rsi("city")%></OPTION>
                                     <% rsi.movenext
    loop
	end if
rsi.close
set rsi = nothing%>
                                   </SELECT>

修改期号：第<input type="text" name="issue" value="<%=issue%>" size="10" <%=onKeyUp%>>期&nbsp;&nbsp;&nbsp;

发行时间：<input type="text" name="Releasetime" size="16" value="<%=Releasetime%>">
&nbsp;&nbsp;&nbsp;状态：<SELECT name="state" id="state">
<OPTION value="true" <%if state="True" then%>selected<%end if%>>开启</OPTION>
<OPTION value="false" <%if state="False" then%>selected<%end if%>>关闭</OPTION></SELECT>
<input type="submit" onclick="javascript:return CheckForm();" value="确定修改" name="B3"></td>
</tr>
</table>
</form>
<%else%>
<form name="formadd" method="POST" action="?action=add"><div align="center">
	
<table width="98%" border="0" class="table" cellspacing="1" cellpadding="0">
<tr> 
<td height="28" bgcolor="#799AE1" class="tdbar">
&nbsp;地方版增加</td>
</tr>
<tr> 
<td align="center" class="td">
<p>（<a href="../admin/city_Level1.asp">管理地区</a>） &nbsp&nbsp;&nbsp;
所属地区：<%Dim rsdq
	  set rsdq=server.createobject("adodb.recordset")
set rsdq=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")
%>
<SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
		//dim count:
		dim count:count = 0
        do while not rsdq.eof 
        %>
subcat[<%=count%>] = new Array("<%=rsdq("city")%>","<%=rsdq("id")%>","<%=rsdq("twoid")%>");
        <%
        count = count + 1
        rsdq.movenext
        loop
        rsdq.close
		set rsdq=nothing
        %>
onecount=<%=count%>;
</SCRIPT>
<%
set rsdq=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid>0")
%>
 <SCRIPT language = "JavaScript">
var onecount4;
onecount4=0;
subcat4 = new Array();
        <%
		dim count4:count4 = 0
        do while not rsdq.eof 
        %>
subcat4[<%=count4%>] = new Array("<%=rsdq("city")%>","<%=rsdq("id")%>","<%=rsdq("twoid")%>","<%=rsdq("threeid")%>");
        <%
        count4 = count4 + 1
        rsdq.movenext
        loop
        rsdq.close
		set rsdq = nothing
        %>
onecount4=<%=count4%>;

function changelocation(locationid)
    {
    document.formadd.city_twoid.length = 0; 
	document.formadd.city_twoid.options[0] = new Option('选择城市','');
	document.formadd.city_threeid.length = 0; 
	document.formadd.city_threeid.options[0] = new Option('选择城市','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.formadd.city_twoid.options[document.formadd.city_twoid.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.formadd.city_threeid.length = 0; 
    document.formadd.city_threeid.options[0] = new Option('选择城市','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.formadd.city_threeid.options[document.formadd.city_threeid.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
      <SELECT name="city_oneid" size="1" id="select2" onChange="changelocation(document.formadd.city_oneid.options[document.formadd.city_oneid.selectedIndex].value)">
        <OPTION value="" selected>选择城市</OPTION>
        <%set rsdq=conn.execute("select * from DNJY_city where id>0 and twoid=0 order by indexid asc")
if rsdq.eof or rsdq.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsdq.eof
response.write "<option value='"&rsdq("id")&"'>"&rsdq("city")&"</option>"
rsdq.movenext
loop
end if
rsdq.close
set rsdq = nothing
%>
      </SELECT>
      <SELECT name="city_twoid" id="city_twoid" onChange="changelocation4(document.formadd.city_twoid.options[document.formadd.city_twoid.selectedIndex].value,document.formadd.city_oneid.options[document.formadd.city_oneid.selectedIndex].value)">
    <OPTION value="" selected>选择城市</OPTION>
	</SELECT>
	     <SELECT name="city_threeid" id="city_threeid">
        <OPTION value="" selected>选择城市</OPTION>
    </SELECT>
&nbsp;&nbsp;&nbsp;新增期号：第<input type="text" name="issue" value="<%=FormatDate(now(),16)%>" size="10" <%=onKeyUp%>>期&nbsp;&nbsp;&nbsp;  
发行时间：<input type="text" name="Releasetime" size="15" value="<%=formatdatetime(now(),1)%>">&nbsp;&nbsp;&nbsp;本期状态：<SELECT name="state" id="state">
<OPTION value="true">开启</OPTION>
<OPTION value="false">关闭</OPTION></SELECT>
&nbsp;&nbsp;<input type="submit" onclick="javascript:return CheckForm();" value="确定添加" name="B2"></td>
</tr>
</table>
</form>
<%end If
Dim k,selectkey
k=1
selectkey=trim(request.form("selectkey"))
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_city] where magazine="&DN_True&""
If selectkey<>"" Then
sql=sql&" and city like '%"&selectkey&"%'"
End if
sql=sql&" order by id"&DN_OrderType&"" 
rs.open sql,conn,1,1
if rs.eof and rs.bof Then
response.write "没有开办杂志的城市"
else%>
<center>
<table width="98%" style="border-collapse: collapse" cellpadding="0" border="0" class="table"><tr>
	<td height="20"  valign="bottom" bgcolor="#BDBEDE" style="border-bottom-style: none; border-bottom-width: medium" class="tdbar" width="67"><p align="center">编号</td>
	<td  width="250" height="20"  valign="bottom" bgcolor="#BDBEDE" style="border-bottom-style: none; border-bottom-width: medium" class="tdbar">开办有电子杂志的地区</td><td  width="150" height="20"  valign="bottom" bgcolor="#BDBEDE" style="border-bottom-style: none; border-bottom-width: medium" class="tdbar">发行情况</td> <td  width="100" height="20"  valign="bottom" bgcolor="#BDBEDE" style="border-bottom-style: none; border-bottom-width: medium" class="tdbar">当前状态</td> <td  width="200" height="20" valign="bottom" bgcolor="#BDBEDE" style="border-bottom-style: none; border-bottom-width: medium" class="tdbar"><p align="center">管理操作</td>
	</tr><%Dim rs_dq,i,j,page,pages
j=0
rs.pagesize=15
page=clng(request("page"))
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.absolutepage=page  

	do while not rs.eof And j<rs.pagesize
    set rs_dq=conn.execute("select count(id) from DNJY_zz_issue where city_oneid="&rs("id")&" and city_twoid="&rs("twoid")&" and city_threeid="&rs("threeid")&"")
	%> <tr>
		<td style="border-bottom:1px solid #C0C0C0; border-left-style:none; border-left-width:1px; border-top-width:1px; border-right-width:1px" height="30" width="69"><div align="center"><b><%=k%></b></div></td>
		<td style="border-bottom:1px solid #C0C0C0; border-top-width:1px; border-left-width:1px; border-right-width:1px" height="20" width="250"><p align="left"><%=conn.Execute("Select city From DNJY_city Where id="&rs("id")&" And twoid=0 And threeid=0")(0)%> <%If rs("twoid")>0 Then response.write conn.Execute("Select city From DNJY_city Where id="&rs("id")&" And twoid="&rs("twoid")&" And threeid=0")(0)%> <%If rs("threeid")>0 Then response.write conn.Execute("Select city From DNJY_city Where id="&rs("id")&" And twoid="&rs("twoid")&" And threeid="&rs("threeid")&"")(0)%></td><td style="border-bottom:1px solid #C0C0C0; border-top-width:1px; border-left-width:1px; border-right-width:1px" height="20" width="150">本地区已发行 <b><%=rs_dq(0)%></b> 期</td><td style="border-bottom:1px solid #C0C0C0; border-top-width:1px; border-left-width:1px; border-right-width:1px" height="20" width="100"><%If rs("state")=true Then response.write "<a href='?action=stateno&city_oneid="&rs("id")&"&city_twoid="&rs("twoid")&"&city_threeid="&rs("threeid")&"'>启用↓</a>" Else response.write "<a href='?action=stateok&city_oneid="&rs("id")&"&city_twoid="&rs("twoid")&"&city_threeid="&rs("threeid")&"'>暂停↓</a>" End if%>&nbsp;&nbsp;&nbsp;<td style="border-bottom:1px solid #C0C0C0; border-top-width:1px; border-left-width:1px; border-right-width:1px" height="20" width="350"><span id="followImg<%=k%>" style="CURSOR: hand" onclick="loadThreadFollow(<%=k%>,5)">控制面板↓</span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="magazine_edition.asp?city_oneid=<%=rs("id")%>&city_twoid=<%=rs("twoid")%>&city_threeid=<%=rs("threeid")%>"><font color="#666666">[<font color="#0000FF">管理此地区期刊</font>]</font></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="?action=alldel&city_oneid=<%=rs("id")%>&city_twoid=<%=rs("twoid")%>&city_threeid=<%=rs("threeid")%>" onClick='return confirm("删除本地区杂志并删除属下原有期刊和相关图片等，操作无法恢复，请确认！");'><font color="#666666">[<font color="#FF0000">删除本地区杂志</font>]</font></a></td>
		</tr><tr style="display:none" id="follow<%=k%>">
			<td height="20"  valign="bottom" colspan="5" bgcolor="#ffffff"><div ><center>
				<table cellpadding=0 width="100%" border="0" style="border-collapse: collapse" bordercolor="#75BB2C" >
				<%i=1
               set rs_b = Server.CreateObject("ADODB.RecordSet")
               sql_b="select * from [DNJY_zz_issue] where city_oneid="&rs("id")&" and city_twoid="&rs("twoid")&" and city_threeid="&rs("threeid")&" order by issue"&DN_OrderType&",id"&DN_OrderType&""
               rs_b.open sql_b,conn,1,1
               if rs_b.eof then
               response.write "<tr><td><font color=""#ff0000""><p align=""center"">没有发行期刊,请在上方添加！</font></td></tr>"
               else
               %> <tr><%do while not rs_b.eof%> <td height=20 width="68"><p align="center"><font color="#666666"><%=i%></font></td>
						<td height=20 width="150"><font color="#666666">第<%=rs_b("issue")%>期</font></td>
						<td height=20 width="150"><font color="#666666">发行时间：<%=rs_b("Releasetime")%></font></td>
						<td height=20 width="200"><font color="#666666">操作时间：<%=rs_b("addtime")%></font></td>
						<td height=20 width="100"><font color="#666666">点击量：<%=rs_b("Click")%>次</font></td>
						<td height=20 width="100"><font color="#666666">本期状态：<%If rs_b("state")=true Then response.write "启用" Else response.write "关闭" End if%></font></td>
						<td height=20 width="80"><font color="#800000"><span style="text-decoration: none"><a href="?id=<%=rs_b("id")%>&action=edit&state=<%=rs_b("state")%>"><font color="#666666">[修改]</font></a></span></font></td>
						<td height=20 width="80">&nbsp;&nbsp;&nbsp;&nbsp;<font color="#800000"><span style="text-decoration: none"><a href="?id=<%=rs_b("id")%>&action=del" onClick='return confirm("确定要删除吗?");'><font color="#666666">[删除]</font></a></span></font></td>
						</tr><%
               i=i+1
			   rs_b.movenext
			   loop
               end if
               rs_b.close
               set rs_b=nothing
               %> </table></center></div></td></tr><%
             k=k+1
			 j=j+1
			 rs.movenext
			 loop
			 %> </table>
			 
	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr bgcolor="#ffffff"> 
<form method=post action="magazine_issue.asp">  
      <td height="30" align="center"> 
    <%if page<2 then      
    response.write "首页 上一页&nbsp;"
  else
    response.write "<a href=magazine_issue.asp?page=1&city_oneid="&city_oneid&"&city_twoid="&city_twoid&"&city_threeid="&city_threeid&">首页</a>&nbsp;"
    response.write "<a href=magazine_issue.asp?page=" & page-1 & "&city_oneid="&city_oneid&"&city_twoid="&city_twoid&"&city_threeid="&city_threeid&">上一页</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "下一页 尾页"
  else
    response.write "<a href=magazine_issue.asp?page=" & (page+1) & "&city_oneid="&city_oneid&"&city_twoid="&city_twoid&"&city_threeid="&city_threeid&">"
    response.write "下一页</a> <a href=magazine_issue.asp?page="&rs.pagecount&"&city_oneid="&city_oneid&"&city_twoid="&city_twoid&"&city_threeid="&city_threeid&">尾页</a>"
  end if
   response.write "&nbsp;页次：<strong><font color=red>"&page&"</font>/"&rs.pagecount&"</strong>页 "
    response.write "&nbsp;共<b><font color='#ff0000'>"&rs.recordcount&"</font></b>条记录 <b>"&rs.pagesize&"</b>条记录/页"
   response.write " 转到：<input type='text' name='page' size=4 maxlength=10 class=input value="&page&">"
   response.write " <input class=input type='submit'  value=' goto '  name='cndok'></span></p>"     
%>
      </td></form>
  </tr>
</table>		 
<%End If
Call CloseDB(conn)%>

    <form name="form_s" method="post" action="magazine_issue.asp">
     
	  <table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
          <tr bgcolor="#FFFFFF" align="center"> 
            <td >地区关键词 <input name="selectkey" type="text" id="selectkey" onFocus="this.value=''" value="请输入地区关健字" size="30"> 
             <input type="submit" name="Submit" value="搜 索" ></td>
          </tr>
        </table>
    </form>
		 
			 </center>
</div></center>