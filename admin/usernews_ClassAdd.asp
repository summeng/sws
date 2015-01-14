<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=../Include/Function.asp-->
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
function checkBig()
{
  if (document.form1.ClassName.value=="")
  {
    alert("类别名称不能为空！");
    document.form1.ClassName.focus();
    return false;
  }
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>添加类别</title>
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
dim Action,ClassName
Action=trim(Request("Action"))
ClassName=trim(request("ClassName"))
if Action="Add" then
   		if ClassName="" then
			response.Write "<script language=javascript>alert( '类别名称不能为空！'  );location.href = 'javascript:history.back()'</script>"
			response.End
		end if
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.open "Select * From DNJY_userNewsClass Where ClassName='" & ClassName & "'",conn,1,3
		if not (rs.bof and rs.EOF) then
			response.Write "<script language=javascript>alert( '类别名称已经存在！'  );location.href = 'javascript:history.back()'</script>"
			response.End
		else
    	 	rs.addnew
     		rs("ClassName")=ClassName    	 	
			rs("showid")=rs("id")
			rs.update
     		rs.Close
	     	set rs=Nothing
    	 	conn.close
			set conn=nothing
			Response.Redirect "usernews_ClassManage.asp"  
		end if
		
		
else
%>
<body>
<div align="center"><br>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">
  <TR> 
    <TD height="5" colspan="12" bgcolor="#799AE1" align="center"><a href="usernews_ClassManage.asp"><strong><font color="#FF0000"><u>类别管理</u></font></strong></a></TD>
  </TR>
    </p>
      <form name="form1" method="post" action="usernews_ClassAdd.asp" onSubmit="return checkBig()">
  <table border="0" cellspacing="1" align="center" class="list1">
    <tr>
      <th colspan="2" ><div align="center">添加类别</div></th>
    </tr>
    <tr>
      <td ><div align="right">类别名称：</div></td>
      <td ><input name="ClassName" type="text" size="20" maxlength="30"></td>
    </tr>
    <tr>
      <td >&nbsp;</td>
      <td height="22" align="center" ><div align="left">
          <input name="Action" type="hidden" id="Action" value="Add">
          <input  style="height:20px;" name="Add" type="submit" class="button" value=" 添 加 ">
      </div></td>
    </tr>
  </table>
  
  </form>
  <p><br>
    <br>
  </p>
</div>
</body>
<%end if%>
</html>
