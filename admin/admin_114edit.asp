<!--#include file="../conn.asp"-->
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
<%call checkmanage("12")
Call OpenConn
dim action,d_name,co,d_tel,dd,hfje,xsts,updatetime,dqsj,admin114
id=strint(request.form("id"))
If request.form("id")="" Then id=strint(request("id"))
action=HtmlEncode(request.querystring("action"))
d_name=HtmlEncode(request("d_name"))
d_tel=HtmlEncode(request("d_tel"))
xsts=trim(request("xsts"))
hfje=trim(request("hfje"))
admin114=trim(request("admin114"))

city_oneid=strint(request.form("city_one"))
city_twoid=strint(request.form("city_two"))
city_threeid=strint(request.form("city_three"))
if city_oneid>0 then
set rs=server.createobject("adodb.recordset")
rs.open "select city from DNJY_city where twoid=0 and threeid=0 and id="&city_oneid,conn,1,1
city_one=rs("city")
rs.close
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
set rs=nothing
end If
select case action
case "add" 
set rs=server.CreateObject("adodb.recordset")
rs.open "select D_name,d_tel,co,city_oneid,city_twoid,city_threeid,city_one,city_two,city_three,hfje,xsts,updatetime,dqsj from DNJY_114",conn,1,3
rs.AddNew
rs(0)=d_name
rs(1)=d_tel
if city_oneid="" then city_oneid=0
if city_twoid="" then city_twoid=0
if city_threeid="" then city_threeid=0
rs("city_oneid")=city_oneid
rs("city_one")=city_one
rs("city_two")=city_two
rs("city_twoid")=city_twoid
rs("city_three")=city_three
rs("city_threeid")=city_threeid
rs("hfje")=trim(request("hfje"))
rs("xsts")=xsts
rs("co")=trim(request("co"))
rs("updatetime")=now()
rs("dqsj")= dateadd("d",xsts, now)
rs.Update
response.Write "<script language=javascript>alert('添加成功!');location.replace('admin_114edit.asp?admin114=2');</script>"
rs.Close:set rs=nothing

case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.open "select d_name,d_tel,hfje,xsts,updatetime,dqsj,co,city_oneid,city_twoid,city_threeid,city_one,city_two,city_three from DNJY_114 where id="&id,conn,1,3
rs(0)=d_name
rs(1)=d_tel
rs("hfje")=trim(request("hfje"))
rs("xsts")=xsts
rs("updatetime")=now()
rs("dqsj")= dateadd("d",xsts, now)
if city_oneid="" then city_oneid=0
if city_twoid="" then city_twoid=0
if city_threeid="" then city_threeid=0
rs("city_oneid")=city_oneid
rs("city_one")=city_one
rs("city_two")=city_two
rs("city_twoid")=city_twoid
rs("city_three")=city_three
rs("city_threeid")=city_threeid
rs("co")=trim(request("co"))
rs.Update
rs.Close:set rs=Nothing
'response.Write "<script language=javascript>alert('修改成功!');location.replace('admin_114edit.asp?admin114=1&id="&id&"');</script>"
response.Write "<script language=javascript>alert('修改成功!');window.close();</script>"
case "del"
Dim Num
if action="del" then
If Request.Form("DelID")="" or isnull(Request.Form("DelID")) or isempty(Request.Form("DelID")) then
response.write "<font style='color:red;font-size:12px'>请选择要删除的信息</font>"
else
Num=request.form("DelID").count
for i=1 to Num
set rs = server.CreateObject ("Adodb.recordset")
rs.open "select * from DNJY_114 where id="&request.form("DelID")(i),conn,3,3
do while not rs.eof
rs.delete
rs.movenext
Loop
Next
Call CloseRs(rs)
end If
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('成功删除！');" & chr(13)
		response.write "window.document.location.href='admin_114.asp';"&chr(13)
		response.write "</script>" & chr(13)
response.end
end If

end select
%>
<script language = "JavaScript">
function CheckForm()
{

	if (document.form.d_name.value.length == 0) {
		alert("请填服务名称！");
		document.form.d_name.focus();
		return false;
	}

		if (document.form.d_tel.value.length == 0) {
		alert("请填服务电话！");
		document.form.d_tel.focus();
		return false;
	}
	if (document.form.hfje.value.length == 0) {
		alert("请填座位票价！");
		document.form.hfje.focus();
		return false;
	}

		if (document.form.xsts.value.length == 0) {
		alert("请填显示期限！");
		document.form.xsts.focus();
		return false;
	}
  

	return true;
}
</script>
<HTML>
<HEAD>
<TITLE>都市114管理</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK rel="stylesheet" href="../css/style.css" type="text/css">
</HEAD>
<BODY>
<TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
  <TR> 
    <TD height="20" align="center" bgcolor="#799AE1"><FONT color="#FFFFFF" style="font-size:14px">&nbsp;
    </FONT><FONT color="#FFFFFF" style="font-size:14px">都 市 114  管 理</FONT></TD>
  </TR>

 <%If admin114="1" Then
 set rs=server.CreateObject("adodb.recordset")
  rs.Open "select d_name,d_tel,id,hfje,xsts,updatetime,dqsj,co,city_oneid,city_twoid,city_threeid from DNJY_114 where id="&id,conn,1,1%>
		  <FORM name="form" method="post" action="?action=edit&id=<%=int(rs("id"))%>&city_oneid=<%=city_oneid%>&city_twoid=<%=city_twoid%>&city_threeid=<%=city_threeid%>" language="javascript" onsubmit="return CheckForm()">
		   <TR> 
    <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">修 改 都 市  114</FONT></TD>
  </TR>
  <TR> 
    <TD bgcolor="#FFFFFF"> <BR>
	<TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
            <tr> 
      <td width="188" height="25" align="right">所属地区：</td>
      <td colspan="2">
<%Dim rsi
set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")
%>
      <SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
		dim count:count = 0
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
		dim count4:count4 = 0
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
      <SELECT name="city_one" size="1" id="select2" onChange="changelocation(document.form.city_one.options[document.form.city_one.selectedIndex].value)">
            <OPTION value="" selected>选择城市</OPTION>
  <%set rsi=conn.execute("select * from DNJY_city where id>0 and twoid=0 and threeid=0  order by indexid asc")
if rsi.eof or rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
  <OPTION value="<%=rsi("id")%>" <%if rsi("id")=rs("city_oneid") then%>selected<%end if%>><%=rsi("city")%></OPTION>
 <%rsi.movenext
    loop
	%>
	<%end if
rsi.close
set rsi = nothing
%>
      </SELECT> 
	  <SELECT name="city_two" id="city_two" onChange="changelocation4(document.form.city_two.options[document.form.city_two.selectedIndex].value,document.form.city_one.options[document.form.city_one.selectedIndex].value)">
    <OPTION value="" selected>选择城市</OPTION>
   <%
set rsi=conn.execute("select * from DNJY_city where id="&rs("city_oneid")&" and twoid>0 and threeid=0")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
  <OPTION value="<%=rsi("twoid")%>" <%if rsi("twoid")=rs("city_twoid") then%>selected<%end if%>><%=rsi("city")%></OPTION>
 <%rsi.movenext
    loop
	%>
	<%end if
rsi.close
set rsi = nothing
%>
	</SELECT>
	     <SELECT name="city_three" id="city_three">
         <OPTION value="" selected>选择城市</OPTION>
		<%
set rsi=conn.execute("select * from DNJY_city where id="&rs("city_oneid")&" and twoid="&rs("city_twoid")&" and threeid<>0")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
<OPTION value="<%=rsi("threeid")%>" <%if rsi("threeid")=rs("city_threeid") then%>selected<%end if%>><%=rsi("city")%></OPTION>
   <% rsi.movenext
    loop
	%>
<%end if
rsi.close
set rsi = nothing
%>
    </SELECT></td>
    </tr>

  <input type="hidden" name="id" id="id" value="<%=rs(2)%>">
    <tr> 
      <td width="188" height="25" align="right">服务名称：</td>
      <td colspan="2">
<INPUT name="d_name" type="text" id="d_name" value="<%=trim(rs(0))%>" size="12"> <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td width="188" height="25" align="right">电&nbsp;&nbsp;&nbsp;&nbsp;话：</td>
      <td colspan="2">
<INPUT name="d_tel" type="text" id="d_tel" value="<%=trim(rs(1))%>" size="20"> <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td width="188" height="25" align="right">座位票价：</td>
      <td colspan="2">
<input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="hfje" size="5" value="<%=rs(3)%>" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"> (元) <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td width="188" height="25" align="right">显示期限：</td>
      <td colspan="2">
<input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="xsts" size="5" value="<%=rs(4)%>" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"> (天) <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td width="188" height="25" align="right">显示颜色：</td>
      <td colspan="2">
<font face="宋体">
<%co=rs("co")%>        
          <select name="co" class="inputa">
            <option style="COLOR: black" value="0" selected>不使用 </option>
            <option style="background: #000088" value="000088" <% if co="000088" Then Response.write("Selected") %>> </option>
            <option style="background: #0000ff" value="0000ff" <% if co="0000ff" Then Response.write("Selected") %>> </option>
            <option style="background: #008800" value="008800" <% if co="008800" Then Response.write("Selected") %>> </option>
            <option style="background: #008888" value="008888" <% if co="008888" Then Response.write("Selected") %>> </option>
            <option style="background: #0088ff" value="0088ff" <% if co="0088ff" Then Response.write("Selected") %>> </option>
            <option style="background: #00a010" value="00a010" <% if co="00a010" Then Response.write("Selected") %>> </option>
            <option style="background: #1100ff" value="1100ff" <% if co="1100ff" Then Response.write("Selected") %>> </option>
            <option style="background: #111111" value="111111" <% if co="111111" Then Response.write("Selected") %>> </option>
            <option style="background: #333333" value="333333" <% if co="333333" Then Response.write("Selected") %>> </option>
            <option style="background: #50b000" value="50b000" <% if co="50b000" Then Response.write("Selected") %>> </option>
            <option style="background: #880000" value="880000" <% if co="880000" Then Response.write("Selected") %>> </option>
            <option style="background: #8800ff" value="8800ff" <% if co="8800ff" Then Response.write("Selected") %>> </option>
            <option style="background: #888800" value="888800" <% if co="888800" Then Response.write("Selected") %>> </option>
            <option style="background: #888888" value="888888" <% if co="888888" Then Response.write("Selected") %>> </option>
            <option style="background: #8888ff" value="8888ff" <% if co="8888ff" Then Response.write("Selected") %>> </option>
            <option style="background: #aa00cc" value="aa00cc" <% if co="aa00cc" Then Response.write("Selected") %>> </option>
            <option style="background: #aaaa00" value="aaaa00" <% if co="aaaa00" Then Response.write("Selected") %>> </option>
            <option style="background: #ccaa00" value="ccaa00" <% if co="ccaa00" Then Response.write("Selected") %>> </option>
            <option style="background: #ff0000" value="ff0000" <% if co="ff0000" Then Response.write("Selected") %>> </option>
            <option style="background: #ff0088" value="ff0088" <% if co="ff0088" Then Response.write("Selected") %>> </option>
            <option style="background: #ff00ff" value="ff00ff" <% if co="ff00ff" Then Response.write("Selected") %>> </option>
            <option style="background: #ff8800" value="ff8800" <% if co="ff8800" Then Response.write("Selected") %>> </option>
            <option style="background: #ff0005" value="ff0005" <% if co="ff0005" Then Response.write("Selected") %>> </option>
            <option style="background: #ff88ff" value="ff88ff" <% if co="ff88ff" Then Response.write("Selected") %>> </option>
            <option style="background: #ee0005" value="ee0005" <% if co="ee0005" Then Response.write("Selected") %>> </option>
            <option style="background: #ee01ff" value="ee01ff" <% if co="ee01ff" Then Response.write("Selected") %>> </option>
            <option style="background: #3388aa" value="3388aa" <% if co="3388aa" Then Response.write("Selected") %>> </option>
            <option style="background: #000000" value="000000" <% if co="000000" Then Response.write("Selected") %>> </option>
          </select>
         
        </font></td>
    </tr>
  <tr bgcolor="#eeeeee">
    <td height="35" colspan="10" align="center"><table width="50%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><div align="center">		
            <input type="submit" value="确认" name="B1" onClick="if(d_name.value=='' || d_tel.value==''){alert('名称和电话都不能为空!');return false;}">
        </div></td>
        <td><div align="center">
            <input type="reset" value="取消" name="B1">
        </div></td>
      </tr>          
        </FORM>
        <%Call CloseRs(rs) %>
      </TABLE>
	<BR></TD>
  </TR>
</TABLE>
<%End if%>
<%If admin114="2" then%>
<FORM name="form" method="post" action="?action=add" language="javascript" onsubmit="return CheckForm()">
<TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
  <TR> 
    <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">添 加 都 市  114</FONT></TD>
  </TR>
  <TR> 
    <TD bgcolor="#FFFFFF">
	<BR>
	<TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
   <TR> 
    <TD bgcolor="#FFFFFF"> <BR>
	<TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
            <tr> 
      <td width="188" height="25" align="right">所属地区：</td>
      <td colspan="2">
		<%
set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")
%>
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
		//dim count4:
		count4 = 0
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
          </SELECT></td>
    </tr>

  
    <tr> 
      <td width="188" height="25" align="right">服务名称：</td>
      <td colspan="2">
<INPUT name="d_name" type="text" id="d_name" value="" size="12"> <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td width="188" height="25" align="right">电&nbsp;&nbsp;&nbsp;&nbsp;话：</td>
      <td colspan="2">
<INPUT name="d_tel" type="text" id="d_tel" value="" size="20"> <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td width="188" height="25" align="right">座位票价：</td>
      <td colspan="2">
<input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="hfje" size="5" value="" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"> (元) <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td width="188" height="25" align="right">显示期限：</td>
      <td colspan="2">
<input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="xsts" size="5" value="" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"> (天) <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td width="188" height="25" align="right">显示颜色：</td>
      <td colspan="2">
<font face="宋体">          
          <select name="co" class="inputa">
            <option style="COLOR: black" value="0" selected>不使用 </option>
            <option style="background: #000088" value="000088"> </option>
            <option style="background: #0000ff" value="0000ff"> </option>
            <option style="background: #008800" value="008800"> </option>
            <option style="background: #008888" value="008888"> </option>
            <option style="background: #0088ff" value="0088ff"> </option>
            <option style="background: #00a010" value="00a010"> </option>
            <option style="background: #1100ff" value="1100ff"> </option>
            <option style="background: #111111" value="111111"> </option>
            <option style="background: #333333" value="333333"> </option>
            <option style="background: #50b000" value="50b000"> </option>
            <option style="background: #880000" value="880000"> </option>
            <option style="background: #8800ff" value="8800ff"> </option>
            <option style="background: #888800" value="888800"> </option>
            <option style="background: #888888" value="888888"> </option>
            <option style="background: #8888ff" value="8888ff"> </option>
            <option style="background: #aa00cc" value="aa00cc"> </option>
            <option style="background: #aaaa00" value="aaaa00"> </option>
            <option style="background: #ccaa00" value="ccaa00"> </option>
            <option style="background: #ff0000" value="ff0000"> </option>
            <option style="background: #ff0088" value="ff0088"> </option>
            <option style="background: #ff00ff" value="ff00ff"> </option>
            <option style="background: #ff8800" value="ff8800"> </option>
            <option style="background: #ff0005" value="ff0005"> </option>
            <option style="background: #ff88ff" value="ff88ff"> </option>
            <option style="background: #ee0005" value="ee0005"> </option>
            <option style="background: #ee01ff" value="ee01ff"> </option>
            <option style="background: #3388aa" value="3388aa"> </option>
            <option style="background: #000000" value="000000"> </option>
          </select>          
        </font></td>
    </tr>
  <tr bgcolor="#eeeeee">
    <td height="35" colspan="10" align="center"><table width="50%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><div align="center">
		
            <input type="submit" value="确认" name="B1" >
        </div></td>
        <td><div align="center">
            <input type="reset" value="取消" name="B1">
        </div></td>
      </tr>          
        </FORM>
        <%Call CloseDB(conn)%>
     
      </TABLE>
	<BR></TD>
  </TR>
</TABLE>
<%End if%>
</BODY> 
</HTML> 
