<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="inc/Article_Function.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
response.buffer=true	
%>

<%
Call OpenConn
dim SqlItem,RsItem,FoundErr,ErrMsg
Dim ItemID,ItemName,WebUrl,city_oneid,city_twoid,city_threeid,city_one,city_two,city_three,c_oneid,c_twoid,c_threeid,c_one,c_two,c_three,LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse,ItemDemo,rso
ItemID=Trim(Request("ItemID"))
If ItemID="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，项目ID不能为空！</li>"
Else
   ItemID=Clng(ItemID)
End If
If FoundErr<>True Then
   SqlItem ="select * from DNJY_Article_Item where ItemID=" & ItemID
   Set RsItem=Server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,Conn,1,1
   If RsItem.Eof And RsItem.Bof  Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，没有找到该项目！</li>"
   Else
      ItemName=RsItem("ItemName")
      ItemDemo=RsItem("ItemDemo")
      WebName=RsItem("WebName")
      WebUrl=RsItem("WebUrl")
	  city_oneid=rsitem("city_oneid")
	  city_twoid=rsitem("city_twoid")
	  city_threeid=rsitem("city_threeid")
	  city_one=rsitem("city_one")
	  city_two=rsitem("city_two")
	  city_three=rsitem("city_three")
    c_oneid=rsitem("c_oneid")
    c_one=rsitem("c_one")
    c_twoid=rsitem("c_twoid")
    c_two=rsitem("c_two")
    c_threeid=rsitem("c_threeid")
    c_three=rsitem("c_three")
      LoginType=RsItem("LoginType")
      LoginUrl=RsItem("LoginUrl")
      LoginPostUrl=RsItem("LoginPostUrl")
      LoginUser=RsItem("LoginUser")
      LoginPass=RsItem("LoginPass")
      LoginFalse=RsItem("LoginFalse")
   End If
   RsItem.Close
   Set RsItem=Nothing
End If

If FoundErr=True Then
   Call WriteErrMsg(ErrMsg)
Else
   Call Main()
End If
'关闭数据库链接
conn.close:Set conn=nothing
%>
<%Sub Main%>
<html>
<head>
<title>采集系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td height="22" colspan="2" align="center" class="topbg"><strong>新闻采集系统项目管理</td>
  </tr>
  <tr class="tdbg"> 
    <td width="65" height="30"><strong>管理导航：</strong></td>
    <td height="30">项目编辑 >> <a href="cj_editor_a.asp?ItemID=<%=ItemID%>"><font color=red>基本设置</font></a> >> <a href="cj_editor_b.asp?ItemID=<%=ItemID%>">列表设置</a> >> <a href="cj_editor_c.asp?ItemID=<%=ItemID%>">链接设置</a> >> <a href="cj_editor_d.asp?ItemID=<%=ItemID%>">正文设置</a> >>  
	<a href="cj_editor_e.asp?ItemID=<%=ItemID%>">采样测试</a> >> <a href="cj_attribute.asp?ItemID=<%=ItemID%>">属性设置</a> >> 完成</td>         
  </tr>         
</table>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
<form method="post" action="cj_editor_b.asp" name="myform">
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>编 辑 项 目--基 本 设 置</strong></div></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" align="right"><strong>项目名称：</strong></td>
      <td class="tdbg" width="75%">
		<input name="ItemName" type="text" size="27" maxlength="30" value="<%=ItemName%>">&nbsp;&nbsp;<font color=red>*</font>如：新浪网－新闻中心 既简单又明了      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" align="right"><strong> 所属地区：</strong></td>
      <td class="tdbg" width="75%"><%
Dim rsi
set rsi=server.createobject("adodb.recordset")
set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")
%>
                                   <SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
		dim count:
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
    document.myform.city_two.length = 0; 
	document.myform.city_two.options[0] = new Option('选择城市','');
	document.myform.city_three.length = 0; 
	document.myform.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.city_two.options[document.myform.city_two.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.myform.city_three.length = 0; 
    document.myform.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.myform.city_three.options[document.myform.city_three.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
                             </SCRIPT>
                                   <SELECT name="city_one" size="1" id="select2" onChange="changelocation(document.myform.city_one.options[document.myform.city_one.selectedIndex].value)">
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
                                   <SELECT name="city_two" id="city_two" onChange="changelocation4(document.myform.city_two.options[document.myform.city_two.selectedIndex].value,document.myform.city_one.options[document.myform.city_one.selectedIndex].value)">
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
                                   <SELECT name="city_three" id="city_three">
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
        <FONT color="#FF0000">*</FONT></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" align="right"><strong> 所属栏目：</strong></td>
      <td class="tdbg" width="75%">        
<%set rsi=conn.execute("select * from DNJY_news_class where oneid>0 and twoid>0 and threeid=0")%>
 <SCRIPT language = "JavaScript">
var onecount2;
onecount2=0;
subcat2 = new Array();
        <%dim count2:count2 = 0
        do while not rsi.eof%>
subcat2[<%=count2%>] = new Array("<%=rsi("name")%>","<%=rsi("oneid")%>","<%=rsi("twoid")%>");
        <%count2 = count2 + 1
        rsi.movenext
        loop
        rsi.close
	set rsi = nothing%>
onecount2=<%=count2%>;
</SCRIPT>
<%set rsi=conn.execute("select * from DNJY_news_class where oneid>0 and twoid>0 and threeid>0")%>
<SCRIPT language = "JavaScript">
var onecount3;
onecount3=0;
subcat3 = new Array();
        <%dim count3:count3 = 0
        do while not rsi.eof%>
subcat3[<%=count3%>] = new Array("<%=rsi("name")%>","<%=rsi("oneid")%>","<%=rsi("twoid")%>","<%=rsi("threeid")%>");
        <%count3 = count3 + 1
        rsi.movenext
        loop
        rsi.close
	set rsi = nothing%>
        onecount3=<%=count3%>;
function changelocation2(locationid)
    {
    document.myform.c_twoid.length = 0; 
    document.myform.c_twoid.options[0] = new Option('选择栏目','');
	document.myform.c_threeid.length = 0; 
    document.myform.c_threeid.options[0] = new Option('选择栏目','');
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
    document.myform.c_threeid.options[0] = new Option('选择栏目','');
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
                                 <SELECT name="c_oneid" size="1" id="select3" onChange="changelocation2(document.myform.c_oneid.options[document.myform.c_oneid.selectedIndex].value)">
                                   <OPTION value="" selected>选择栏目</OPTION>
<%set rsi=conn.execute("select * from DNJY_news_class where oneid>0 and twoid=0 order by sorting asc")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
                                   <OPTION value="<%=rsi("oneid")%>" <%if rsi("oneid")=c_oneid then%>selected<%end if%>><%=rsi("name")%></OPTION>
<% rsi.movenext
    loop
	end if
rsi.close
set rsi= nothing%>
                                 </SELECT>
                                 <SELECT name="c_twoid" id="select4" onChange="changelocation3(document.myform.c_twoid.options[document.myform.c_twoid.selectedIndex].value,document.myform.c_oneid.options[document.myform.c_oneid.selectedIndex].value)">
                                   <OPTION value="" selected>选择栏目</OPTION>
<%set rsi=conn.execute("select * from DNJY_news_class where oneid="&c_oneid&" and twoid<>0 and threeid=0")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
                                   <OPTION value="<%=rsi("twoid")%>" <%if rsi("twoid")=c_twoid then%>selected<%end if%>><%=rsi("name")%></OPTION>
<% rsi.movenext
    loop
	end if
rsi.close
set rsi= nothing%>
                                 </SELECT>
                                 <SELECT name="c_threeid" id="select5">
                                   <OPTION value="" selected>选择栏目</OPTION>
                                   <%set rsi=conn.execute("select * from DNJY_news_class where oneid="&c_oneid&" and twoid="&c_twoid&" and threeid<>0")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
                                   <OPTION value="<%=rsi("threeid")%>" <%if rsi("threeid")=c_threeid then%>selected<%end if%>><%=rsi("name")%></OPTION>
<% rsi.movenext
    loop
	end if
rsi.close
set rsi= Nothing%>
</SELECT>
 </td>
    </tr>
	                       
                  
    <tr>
      <td width="20%" class="tdbg" align="right"><strong> 网站名称：</strong></td>
      <td class="tdbg" width="75%">
		<input name="WebName" type="text" size="27" maxlength="30" value="<%=WebName%>">      </td>
    </tr>
    <tr>
      <td width="20%" class="tdbg" align="right"><strong> 网站网址：</strong></td>
      <td class="tdbg" width="75%"><input name="WebUrl" type="text" size="49" maxlength="100" value="<%=WebUrl%>">      </td>
    </tr>
   <tr class="tdbg"> 
      <td width="20%" class="tdbg" align="right"><strong> 网站登录：</strong></td>
      <td class="tdbg">
		<input type="radio" value="0" name="LoginType" <%if LoginType=0 Then Response.Write "checked"%> onClick="Login.style.display='none'">不需要登录<span lang="en-us">&nbsp;
		</span>
		<input type="radio" value="1" name="LoginType" <%if LoginType=1 Then Response.Write "checked"%> onClick="Login.style.display=''">设置参数      </td>
    </tr>
   <tr class="tdbg" id="Login" style="<%If LoginType=0 Then Response.write "display:none"%>"> 
      <td width="20%" class="tdbg" align="right"><strong> 登录参数：</strong></td>
      <td class="tdbg">
        登录地址：<input name="LoginUrl" type="text" size="40" maxlength="100" value="<%=LoginUrl%>"><br>
        提交地址：<input name="LoginPostUrl" type="text" size="40" maxlength="100" value="<%=LoginPostUrl%>"><br>
        用户参数：<input name="LoginUser" type="text" size="30" maxlength="30" value="<%=LoginUser%>"><br>
        密码参数：<input name="LoginPass" type="text" size="30" maxlength="30" value="<%=LoginPass%>"><br> 
		失败信息：<input name="LoginFalse" type="text" size="30" maxlength="50" value="<%=LoginFalse%>"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" align="right"><strong>模板备注：</strong></td>
      <td class="tdbg" width="75%"><textarea name="ItemDemo" cols="49" rows="5"><%=ItemDemo%></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td colspan="2" align="center" class="tdbg">
        <input name="ItemID" type="hidden" id="ItemID" value="<%=ItemID%>">
        <input name="Action" type="hidden" id="Action" value="SaveEdit">
        <input name="Cancel" type="button" id="Cancel" value=" 返&nbsp;&nbsp;回 " onClick="window.location.href='cj_start.asp'" style="cursor: hand;background-color: #cccccc;">
        &nbsp; 
        <input  type="submit" name="Submit" value="下&nbsp;一&nbsp;步" style="cursor: hand;background-color: #cccccc;"></td>
    </tr>
</form>
</table>
</body>         
</html>
<%End  Sub%>
