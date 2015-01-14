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
call checkmanage("06")
Dim newssc
newssc=15
Call OpenConn%>
<html>
<head>
<title>添加文章</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<script language = "JavaScript">
function CheckForm()
{
	if (document.addNEWS.title.value.length == 0) {
		alert("文章标题没有填写.");
		document.addNEWS.title.focus();
		return false;
	}
	if (document.addNEWS.keywords.value.length == 0) {
		alert("文章关键词没有填写.");
		document.addNEWS.keywords.focus();
		return false;
	}
        var editor=KindEditor.create('#editor');
        if (editor.isEmpty())
	    {
	    alert("文章内容不能为空！")	  
	     return false
	     }
		if (document.addNEWS.newsuser.value.length == 0) {
		alert("文章作者没有填写");
		document.addNEWS.newsuser.focus();
		return false;
	}
		if (document.addNEWS.come.value.length == 0) {
		alert("文章来源没有填写");
		document.addNEWS.come.focus();
		return false;
	}
  
	return true;
}
</script>
<script language=javascript src="../Include/mouse_on_title.js"></script>
<!--kindeditor编辑器-->
	<link rel="stylesheet" href="../KindEditor/themes/default/default.css" />
	<link rel="stylesheet" href="../KindEditor/plugins/code/prettify.css" />
	<script charset="utf-8" src="../KindEditor/kindeditor.js"></script>
	<script charset="utf-8" src="../KindEditor/lang/zh_CN.js"></script>
	<script charset="utf-8" src="../KindEditor/plugins/code/prettify.js"></script>
	<script>
		KindEditor.ready(function(K) {
			var editor = K.create('textarea[name="content"]', {
				cssPath : '../KindEditor/plugins/code/prettify.css',
				uploadJson : '../KindEditor/asp/upload_json.asp?action=upnews',
				afterBlur: function(){this.sync();},//失去焦点执行this获得值
				allowFileManager : true,				
				afterCreate : function() {
					var self = this;
					K.ctrl(document, 13, function() {
						self.sync();
						K('form[name=example]')[0].submit();
					});
					K.ctrl(self.edit.doc, 13, function() {
						self.sync();
						K('form[name=example]')[0].submit();
					});
				}
			});
			prettyPrint();
		});
	</script>
<!--kindeditor编辑器END-->

</head>
<body leftmargin="0" topmargin="0" bgcolor="#FFFFEE">
<form name="addNEWS" method="post" action="admin_article_addsave.asp" onSubmit="return CheckForm()" language="javascript" >
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6699cc">
<tr class=backs><td align="center" class=td height=18 colspan="5"><font color="#FFFFFF">添加网站文章</font></td>
  </tr>
  
    <tr> 
      <td height="30" colspan="3" align="center">文章相关相册在文章添加后批量上传</td>
    </tr>

    <tr> 
      <td width="100" height="25" align="right">&nbsp;&nbsp;所属地区：</td>
      <td colspan="2">
	  <%Dim rsdq
	  set rsdq=server.createobject("adodb.recordset")
set rsdq=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")
%>
<SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
		dim count:
		count = 0
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
    document.addNEWS.city_two.length = 0; 
	document.addNEWS.city_two.options[0] = new Option('选择城市','');
	document.addNEWS.city_three.length = 0; 
	document.addNEWS.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.addNEWS.city_two.options[document.addNEWS.city_two.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.addNEWS.city_three.length = 0; 
    document.addNEWS.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.addNEWS.city_three.options[document.addNEWS.city_three.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
      <SELECT name="city_one" size="1" id="select2" onChange="changelocation(document.addNEWS.city_one.options[document.addNEWS.city_one.selectedIndex].value)">
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
      <SELECT name="city_two" id="city_two" onChange="changelocation4(document.addNEWS.city_two.options[document.addNEWS.city_two.selectedIndex].value,document.addNEWS.city_one.options[document.addNEWS.city_one.selectedIndex].value)">
    <OPTION value="" selected>选择城市</OPTION>
	</SELECT>
	     <SELECT name="city_three" id="city_three">
        <OPTION value="" selected>选择城市</OPTION>
    </SELECT></td>
    </tr>
<tr>
<td width="100" height="25" align="right">&nbsp;&nbsp;所属栏目：</td>
<td colspan="2"><%
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
    document.addNEWS.c_twoid.length = 0; 
    document.addNEWS.c_twoid.options[0] = new Option('二级栏目','');
	document.addNEWS.c_threeid.length = 0; 
    document.addNEWS.c_threeid.options[0] = new Option('三级栏目','');
    var locationid=locationid;
	var i;
    for (i=0;i < onecount2; i++)
        {
            if (subcat2[i][1] == locationid)
            { 
                document.addNEWS.c_twoid.options[document.addNEWS.c_twoid.length] = new Option(subcat2[i][0], subcat2[i][2]);
            }        
        }
       
    }    
     function changelocation3(locationid,locationid1)
    {
    document.addNEWS.c_threeid.length = 0; 
    document.addNEWS.c_threeid.options[0] = new Option('三级栏目','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount3; i++)
        {
            if (subcat3[i][2] == locationid)
            { 
			if (subcat3[i][1] == locationid1)
			{
                document.addNEWS.c_threeid.options[document.addNEWS.c_threeid.length] = new Option(subcat3[i][0], subcat3[i][3]);
            }        
        }
       } 
    }            
	      </SCRIPT>
          <SELECT name="c_oneid" size="1" id="select8" onChange="changelocation2(document.addNEWS.c_oneid.options[document.addNEWS.c_oneid.selectedIndex].value)">
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
          <SELECT name="c_twoid" id="select9" onChange="changelocation3(document.addNEWS.c_twoid.options[document.addNEWS.c_twoid.selectedIndex].value,document.addNEWS.c_oneid.options[document.addNEWS.c_oneid.selectedIndex].value)">
            <OPTION value="" selected>二级栏目</OPTION>
          </SELECT>
          <SELECT name="c_threeid" id="select10">
            <OPTION value="" selected>三级栏目</OPTION>
          </SELECT> <font color="#ff0000"> *</font></td>
                        </tr>
    <tr> 
      <td width="100" height="25" align="right">文章标题：</td>
      <td colspan="2">
<input name="title" type="text" class="input" size="70" maxlength="100"> <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td width="100" height="25" align="right">关 键 词：</td>
      <td colspan="2">
<input name="keywords" type="text" class="input" size="50" maxlength="40" value="">  <font color="#FF0000">*</font> <input type="button" name="btn" value="从标题复制" title="复制标题到关键词" onClick="CopyWebTitle(document.addNEWS.title.value);"> <br>(用英文“,”号分隔，限40个字符)</td>
    </tr>
	<tr>
  <td height="25"  align="right">文章属性：</td>
  <td colspan="2">
          <select class='css_select' name="oStyle" type="text" id="oStyle">
          <option value="s">标题样式</option>
          <option value="s01">粗体</option>
          <option value="s02">斜体</option>
        </select>
        <select class='css_select' name="oColor" type="text" id="oColor">
          <option value="links">标题颜色</option>
          <option value="a01" style="background-color:#FF0000;"></option>
          <option value="a02" style="background-color:#0000FF;"></option>
          <option value="a03" style="background-color:#00FFFF;"></option>
          <option value="a04" style="background-color:#FF9900;"></option>
          <option value="a05" style="background-color:#339966;"></option>
        </select>
        <select class='css_select' name="newstop" type="text" id="newstop">
          <option value="0">置顶</option>
          <option value="1">是</option>
          <option value="0">否</option>
          </select>
		 <select class='css_select' name="tuijian" type="text" id="tuijian">
                <option value="0" selected>推荐</option>
                <option value="1">是</option>
                <option value="0">否</option>
        </select></td>
</tr>
    <tr> 
      <td height="25" valign="top" align="right">文章内容：</td>
      <td colspan="2" valign="top"><textarea id="editor" name="content" style="width:670px;height:400px;"></textarea><input type="checkbox" name="T_SaveImg" value="1" /> 保存内容中远程图片到本地&nbsp;&nbsp;<input type="checkbox" name="AspJpeg" value="1" checked/>给上传图片加水印&nbsp;&nbsp;<font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td height="25" align="right">文章作者：</td>
      <td colspan="2"> 
        <input name="newsuser" type="text" class="input" value="<% = request.cookies("administrator") %>" size="30"> <font color="#FF0000">*</font> &lt;=<%If session("newsuser")<>"" Then response.write "【<FONT color=blue><FONT style=""CURSOR: hand"" onclick=""newsuser.value='"&session("newsuser")&"'"" color=#993300>"&session("newsuser")&"</FONT></FONT>】"%>【<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="newsuser.value='站长'" 
      color=#993300>站长</FONT></FONT>】【<FONT 
      color=blue><FONT style="CURSOR: hand" onclick='newsuser.value="未知"' 
      color=#993300>未知</FONT></FONT>】【<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="newsuser.value='佚名'" 
      color=#993300>佚名</FONT></FONT>】</td>
    </tr>
	    <tr> 
      <td height="25" align="right">文章来源：</td>
      <td colspan="2"> 
        <input name="come" type="text" class="input" value="<%=title%>" size="30" maxlength="30"> <font color="#ff0000">*</font> &lt;=<%If session("come")<>"" Then response.write "【<FONT color=blue><FONT style=""CURSOR: hand"" onclick=""come.value='"&session("come")&"'"" color=#993300>"&session("come")&"</FONT></FONT>】"%>【<FONT color=blue><FONT style="CURSOR: hand" onclick="come.value='因特网'" color=#993300>因特网</FONT></FONT>】</td>
    </tr> 
     <tr> 
      <td height="25" align="right">文章显示：</td>
      <td colspan="2"> 
        <input type="radio" name="ifshow" value="1" checked>
        显示 
        <input name="ifshow" type="radio" value="0" >
        隐藏</td>
    </tr>   
    <tr> 
      <td height="30" colspan="3" align="center"> 
        <input type="submit" name="Submit" value="提交" class="input" >
        <input name="subsave" type="hidden" value="subok">
        <input type="reset" name="Submit2" value="重置" class="input"> 
 </td>
    </tr>
  
  <tr> 
    <td height="30" colspan="3" align="center">&nbsp;</td>
  </tr>
</table></form>
</body>
</html>
