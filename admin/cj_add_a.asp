<!--#include file="../conn.asp"-->
<!--#include file=cookies.asp-->
<!--#include file=../config.asp-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="inc/Article_Function.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%response.buffer=true	

Call Main()
'关闭数据库链接
Conn.Close:Set Conn=Nothing
Sub Main
Call OpenConn%>
<html>
<head>
<title>新闻采集系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td height="22" colspan="2" align="center" class="topbg"><strong>新闻采集系统项目管理</td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr class="tdbg"> 
    <td width="65" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="cj_add_a.asp">添加项目</a> >> <font color=red>基本设置</font> >> 列表设置 >> 链接设置 >> 正文设置 >> 采样测试 >> 属性设置 >> 完成</td>         
  </tr>         
</table>
<br>         
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
<form method="post" action="cj_add_b.asp" name="myform">
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>添 加 新 
		项 目--基 本 设 置</strong></div></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" align="right"><strong>项目名称：</strong></td>
      <td class="tdbg" width="75%">
		<input name="ItemName" type="text" size="27" maxlength="30">&nbsp;&nbsp;<font color=red>*</font>如：新浪网－新闻中心 既简单又明了
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" align="right"><strong> 所属地区：</strong></td>
      <td class="tdbg" width="75%"><%
set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")
%>
        <SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
		dim count:count = 0
        do while not rs.eof 
        %>
subcat[<%=count%>] = new Array("<%=rs("city")%>","<%=rs("id")%>","<%=rs("twoid")%>");
        <%
        count = count + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount=<%=count%>;
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
    document.all.city_two.length = 0; 
	document.all.city_two.options[0] = new Option('选择城市','');
	document.all.city_three.length = 0; 
	document.all.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.all.city_two.options[document.all.city_two.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.all.city_three.length = 0; 
    document.all.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.all.city_three.options[document.all.city_three.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
        <SELECT name="city_one" size="1" id="select2" onChange="changelocation(document.all.city_one.options[document.all.city_one.selectedIndex].value)">
          <OPTION value="" selected>选择城市</OPTION>
          <%set rs=conn.execute("select * from DNJY_city where id>0 and twoid=0 order by indexid asc")
if rs.eof or rs.bof then
response.write "<option value=''>没有分类</option>"
else
do until rs.eof
response.write "<option value='"&rs("id")&"'>"&rs("city")&"</option>"
    rs.movenext
    loop
end if
rs.close
set rs = nothing
%>
        </SELECT>
        <SELECT name="city_two" id="city_two" onChange="changelocation4(document.all.city_two.options[document.all.city_two.selectedIndex].value,document.all.city_one.options[document.all.city_one.selectedIndex].value)">
          <OPTION value="" selected>选择城市</OPTION>
        </SELECT>
        <SELECT name="city_three" id="city_three">
          <OPTION value="" selected>选择城市</OPTION>
        </SELECT>
        <FONT color="#FF0000">*</FONT> </td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" align="right"><strong> 所属栏目：</strong></td>
      <td class="tdbg" width="75%">
<%
set rs=conn.execute("select * from DNJY_news_class where oneid>0 and twoid>0 and threeid=0 order by sorting")
%>
<SCRIPT language = "JavaScript">
var onecount2;
onecount2=0;
subcat2 = new Array();
        <%dim count2:count2 = 0
        do while not rs.eof%>
subcat2[<%=count2%>] = new Array("<%=rs("name")%>","<%=rs("oneid")%>","<%=rs("twoid")%>");
        <%count2 = count2 + 1
        rs.movenext
        loop
Call CloseRs(rs)%>
onecount2=<%=count2%>;
          </SCRIPT>
          <%set rs=conn.execute("select * from DNJY_news_class where oneid>0 and twoid>0 and threeid>0 order by sorting")%>
          <SCRIPT language = "JavaScript">
var onecount3;
onecount3=0;
subcat3 = new Array();
        <%dim count3:count3 = 0
        do while not rs.eof%>
subcat3[<%=count3%>] = new Array("<%=rs("name")%>","<%=rs("oneid")%>","<%=rs("twoid")%>","<%=rs("threeid")%>");
        <%count3 = count3 + 1
        rs.movenext
        loop
Call CloseRs(rs)%>
onecount3=<%=count3%>;

function changelocation2(locationid)
    {
    document.myform.c_twoid.length = 0; 
    document.myform.c_twoid.options[0] = new Option('二级栏目','');
	document.myform.c_threeid.length = 0; 
    document.myform.c_threeid.options[0] = new Option('三级栏目','');
    var locationid=locationid;
	var i;
    for (i=0;i < onecount2; i++)
        {
            if (subcat2[i][1] == locationid)
            { 
                document.myform.c_twoid.options[document.myform.c_twoid.length] = new Option(subcat2[i][0], subcat2[i][2]);
            }        
        }
       
    }    
     function changelocation3(locationid,locationid1)
    {
    document.myform.c_threeid.length = 0; 
    document.myform.c_threeid.options[0] = new Option('三级栏目','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount3; i++)
        {
            if (subcat3[i][2] == locationid)
            { 
			if (subcat3[i][1] == locationid1)
			{
                document.myform.c_threeid.options[document.myform.c_threeid.length] = new Option(subcat3[i][0], subcat3[i][3]);
            }        
        }
       } 
    }            
	      </SCRIPT>
          <SELECT name="c_oneid" size="1" id="select8" onChange="changelocation2(document.myform.c_oneid.options[document.myform.c_oneid.selectedIndex].value)">
            <OPTION value="" selected>一级栏目</OPTION>
            <%set rs=conn.execute("select * from DNJY_news_class where oneid>0 and twoid=0 order by sorting asc")
if rs.eof or rs.bof then
response.write "<option value=''>没有栏目</option>"
else
do until rs.eof
response.write "<option value='"&rs("oneid")&"'>"&rs("name")&"</option>"
    rs.movenext
    loop
end if
Call CloseRs(rs)%>
          </SELECT>
          <SELECT name="c_twoid" id="select9" onChange="changelocation3(document.myform.c_twoid.options[document.myform.c_twoid.selectedIndex].value,document.myform.c_oneid.options[document.myform.c_oneid.selectedIndex].value)">
            <OPTION value="" selected>二级栏目</OPTION>
          </SELECT>
          <SELECT name="c_threeid" id="select10">
            <OPTION value="" selected>三级栏目</OPTION>
          </SELECT></td>
    </tr>
    <tr>
      <td width="20%" class="tdbg" align="right"><strong> 网站名称：</strong></td>
      <td class="tdbg" width="75%">
		<input name="WebName" type="text" size="27" maxlength="30">
      </td>
    </tr>
    <tr>
      <td width="20%" class="tdbg" align="right"><strong> 网站网址：</strong></td>
      <td class="tdbg" width="75%"><input name="WebUrl" type="text" size="49" maxlength="100">
      </td>
    </tr>
   <tr class="tdbg"> 
      <td class="tdbg" align="right"><strong> 网站登录：</strong></td>
      <td class="tdbg">
		<input type="radio" value="0" name="LoginType" checked onClick="Login.style.display='none'">不需要登录<span lang="en-us">&nbsp;
		</span>
		<input type="radio" value="1" name="LoginType" onClick="Login.style.display=''">设置参数</td>
    </tr>
   <tr class="tdbg" id="Login" style="display:none"> 
      <td class="tdbg" align="right" align="right"><strong> 登录参数：</strong></td>
      <td class="tdbg">
        登录地址：<input name="LoginUrl" type="text" size="40" maxlength="100" value=""><br>
        提交地址：<input name="LoginPostUrl" type="text" size="40" maxlength="100" value=""><br>
        用户参数：<input name="LoginUser" type="text" size="30" maxlength="30" value=""><br>
        密码参数：<input name="LoginPass" type="text" size="30" maxlength="30" value=""><br> 
		失败信息：<input name="LoginFalse" type="text" size="30" maxlength="50" value=""></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" align="right"><strong>项目备注：</strong></td>
      <td class="tdbg" width="75%"><textarea name="ItemDemo" cols="49" rows="5"></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td colspan="2" align="center" class="tdbg"><input name="Action" type="hidden" id="Action" value="SaveAdd">
        <input name="Cancel" type="button" id="Cancel" value=" 返&nbsp;&nbsp;回 " onClick="window.location.href='cj_management.asp'" style="cursor: hand;background-color: #cccccc;">
        &nbsp; 
        <input  type="submit" name="Submit" value="下&nbsp;一&nbsp;步" style="cursor: hand;background-color: #cccccc;"></td>
    </tr>
</form>
</table>    
</body>         
</html>
<%end Sub%>
