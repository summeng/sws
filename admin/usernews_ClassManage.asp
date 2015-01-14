<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("04")%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<script language="JavaScript" type="text/JavaScript">

function ConfirmDelBig()
{
   if(confirm("确定要删除此栏目吗？删除此栏目同时将删除该类下的所有记录，并且不能恢复！"))
     return true;
   else
     return false;
	 
}


</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>栏目管理</title>
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
</head>
<%

dim rsClass,rsSmallClass,ErrMsg
set rsClass=server.CreateObject("adodb.recordset")
rsClass.open "Select * From DNJY_userNewsClass",conn,1,3

%>
<body>
<div align="center">
      <br>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">

  <TR> 
    <TD height="5" colspan="12" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">开有店铺的会员文章分类管理</FONT></TD>
  </TR>
      <a href="usernews_ClassAdd.asp"><strong><font color="#FF0000" size="2"><u>添加栏目</u></font></strong></a>
      </p>
      
     <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%" height="30">
        <tr >
          <td width="25%" class="style1"><div align="center">栏目名称</div></td>
          <td width="12%" class="style1"><div align="center">ID号</div></td>
          <td width="13%" class="style1"><div align="center">导航显示序号</div></td>
          <td width="13%" class="style1"><div align="center">导航显示</div></td>
          <!--td width="17%" class="style1"><div align="center">首页栏目栏目显示</div></td-->
          <td width="20%" class="style1"><div align="center">操作选项</div></td>
        </tr>
		
		<%do while not rsClass.eof%>
		 
        <tr>
          <td ><%=rsClass("ClassName")%></td>
          <td ><div align="center"><%=rsClass("ID")%></div></td>
          <td ><div align="center"><%=rsClass("showid")%></div></td>
          <td ><div align="center">
		  <%if rsClass("dhshow")=true then %>
		  		<a href='usernews_classshow.asp?act=dh&sh=0&id=<%=rsClass("id")%>'><font color="#339933">是</font></a> 
	 <% else%>
	  <a href='usernews_classshow.asp?act=dh&sh=1&id=<%=rsClass("id")%>'><font color="#CC0000">否</font></a> 
			<%end if
				
		  %></div></td>
          <!--td >
		  <div align="center">
		  <%'if rsClass("lmshow")=true then %>
		  		<a href='usernews_classshow.asp?act=lm&sh=0&id=<%=rsClass("id")%>'><font color="#339933">是</font></a> 
	 <%' else%>
	  <a href='usernews_classshow.asp?act=lm&sh=1&id=<%=rsClass("id")%>'><font color="#CC0000">否</font></a> 
			<%'end if
				
		  %></div></td-->
          <td ><div align="center"><a href="usernews_ClassModify.asp?ClassID=<%=rsClass("ID")%>">修改</a> | <a href="usernews_ClassDel.asp?ClassID=<%=rsClass("ID")%>"onClick="return ConfirmDelBig();">删除</a>&nbsp;</div></td>
        </tr>
		<%
	  
	  rsClass.movenext
	loop
%>
  </table>
      <table class="list1" width="666" border="0" cellspacing="1" cellpadding="0">
        <tr>
          <td width="666">
            <div align="left">&nbsp;&nbsp;点击“是”或“否”可改变显示状态<br>
            “导航显示序号”是栏目名称在导航栏的排列次序，序号为数字，序号越小越靠前排列
              <br>
              “导航显示”状态“是”在导航栏显示此栏目名称，状态“否”在导航栏不显示此栏目名称<br>
          <!--“首页栏目栏目显示”状态“是”在首页分类栏目中显示此栏目信息，状态“否”在首页分类栏目中不显示此栏目信息--></div></td>
        </tr>
  </table>
      <p>&nbsp;</p>
</div>
<%
rsClass.close
set rsClass=nothing
%>
</table>
</body>
</html>
