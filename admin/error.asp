<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%sub diserror()%>
<HTML>
<HEAD>
<TITLE> New Document </TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"><style type="text/css">
<!--
body {
	background-image: url(../images/dj_bg.gif);
}
-->
</style>
<link href="../css/style.css" rel="stylesheet" type="text/css">
</head>
<BODY><BODY><p>&nbsp;</p><table width="65%" height=180 border="0" cellspacing="0" cellpadding="4" align="center" bordercolor="#000000" style="border-collapse: collapse">
  <tr> 
    <td align="center" bgcolor="#DEDBEF" height=100><%=errmsg%></td>
  </tr>
</table>
</body>
<%end sub%>
<%sub disok()%>
<HTML>
<HEAD>
<TITLE> New Document </TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"<style type="text/css">
<!--
body {
	background-image: url(../images/dj_bg.gif);
}
-->
</style>
<link href="../css/style.css" rel="stylesheet" type="text/css"></head>
<BODY><BODY><p>&nbsp;</p><table width="65%" height=180 border="0" cellspacing="0" cellpadding="4" align="center" bordercolor="#000000" style="border-collapse: collapse">

  <tr> 
    <td align="center" bgcolor="#DEDBEF" height=100><%=errmsg%></td>
  </tr>

</table>
</body>
<%end sub%>