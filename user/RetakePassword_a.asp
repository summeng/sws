<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<meta http-equiv="Content-Language" content="zh-cn">
<link href="/<%=strInstallDir%>css/css.css" rel="stylesheet" type="text/css">
<title>密码找回</title>

<style>
<!--
.style1 {
	font-size: 16px;
	font-family: "黑体";
}
-->
</style>
<script language="javascript">
<!--
if (parent.frames.length > 0) {
parent.location.href = location.href;
}
function form1_onsubmit() {
if (document.form1.username.value=="")
	{
	  alert("用户名都没有，您怎么能找回您的密码呢？")
	  document.form1.username.focus()
	  return false
	 }
	  
}
// --></script>
</head>

<body topmargin="0" leftmargin="0">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="400" id="AutoNumber1" height="147">
    <tr> 
      <td width="500" height="25" bgcolor="#FF9900"> <p align="center">&nbsp; ..:::第一步：请输入你的用户名:::...</td>
    </tr>
  <tr> 
    <td width="500" height="122"> 
      <table border="0" width="100%" id="table1" height="100">
      <form method="POST" name="form1" language="javascript" onSubmit="return form1_onsubmit()" action="RetakePassword_b.asp">
		<tr>
			<td>
			<p align="center"> 输入你的用户帐号（用户名）</td>
		</tr>
		<tr>
			<td height="38">
			<p align="center"><span class="style1">请输入用户名：</span> <input type="text" name="username" size="30" class="form" style="font-family: 宋体; font-size: 9pt; maxlength="20"></td>
		</tr>
		<tr>
			<td>
			<p align="center">
			 <input type="submit" value=" 到下一步 " name="submit" style="font-family: 宋体; font-size: 9pt;>&nbsp;&nbsp;
        <input type="reset" name="Submit2" value="重置">
      </td>
		</tr></form>
		</table>
	</td>
  </tr>
  </table>
<table border="0" width="400" id="table2" bgcolor="#FF9900">
	<tr>
		<td>　</td>
	</tr>
</table>