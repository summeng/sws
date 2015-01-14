<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/tt_auto_lock.asp" -->
<%response.buffer="True"%>
<!--#include file="../Include/md5.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<script language="JavaScript">

<!--
function get_onsubmit() {
if (document.get.userid.value=="")
	{
	  alert("请设定您的登陆名。")
	  document.get.userid.focus()
	  return false
	 }	 	 
}
// -->
</script>
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

</head>
<% dim fpass1,fpass,strSql,username
Call OpenConn
if request("action")="edit" then
username=trim(request.form("username"))
fpass1=trim(request.form("fpass1"))
fpass=md5(fpass1)
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from [DNJY_user] where  username='"&username&"'",conn,1,3
if rs.eof or rs.bof then
response.write "<li>参数错误！"
Call CloseRs(rs)
Call CloseDB(conn)
response.end
end if
rs("password")=fpass
rs.update
Call CloseRs(rs)
Call CloseDB(conn)
response.write "<script>alert('密码修改成功，请记住！');history.back();</Script>"
response.end
end if%>

<%dim answer1,answer
if Request.form("submit")<>"" then 
username=request.form("username") '页面之间传递参数
answer1= trim(Request.form("answer1"))
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_user] where username='"&UserName&"'"
rs.open sql,conn,1,1
answer =rs("answer")
answer1=md5(answer1)
Call CloseRs(rs)
if answer<>answer1 Or answer="d41d8cd98f00b204e9800998ecf8427e" then
session("gpw_error")=session("gpw_error")+1
response.write "<script>alert('你的答案错误，请返回检查！你已错误输入"&session("gpw_error")&"次，超限将锁定IP无法访问！');history.back();</Script>"
Call CloseDB(conn)
response.end
else
end if
end if

Call CloseDB(conn)%>
<center>
<script language="VBScript">
<!--
function checkaddurl()
fpass1=urlform.fpass1.value
if urlform.fpass1.value="" or len(fpass1)<5 or len(fpass1)>12 then 
msgbox"请输入新密码（限5--12个字符）^_^"
urlform.fpass1.focus
elseif urlform.fpass1.value<>"" and urlform.fpass2.value<>urlform.fpass1.value then 
msgbox"确认密码不相同！^_^"
urlform.fpass2.focus
else
urlform.submit
end if

end function
-->
</script>
<body topmargin="0" leftmargin="0">
<%If request.form("username")<>"" then%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="400" id="AutoNumber1" height="147">
    <tr> 
      <td width="500" height="25" bgcolor="#FF9900"> <p align="center">&nbsp; ..:::第三步：请设置你的新密码:::...</td>
    </tr>
  <tr> 
    <td width="500" height="122"> 
      <table border="0" width="100%" id="table1" height="100">
  <form name="urlform" method="post" action="RetakePassword_c.asp?action=edit" onsubmit="return(search4())" >
    
  
    <tr valign="baseline"> 
      <td nowrap align="right"  class=tdc><span class="style1">输入新密码：</span></td>
      <td  class=tdc> 
        <input type="password" name="fpass1" value="" size="20" maxlength="12">（5-12字符内）
      </td>
    </tr>
     <tr valign="baseline"> 
      <td nowrap align="right"  class=tdc><span class="style1">密码确认：</span></td>
      <td  class=tdc> 
        <input type="password" name="fpass2" value="" size="20" maxlength="12">（5-12字符内） 
      </td>
    </tr>
     
        <tr valign="baseline"> 
      <td nowrap align="right"  class=tdc>　</td>
      <td  class=tdc> 
	   <input type="hidden" name="username" value="<%=trim(request.form("username"))%>">
        <input type="button" name="sss" value="提交" onclick=checkaddurl()>
              　 
              <input type="button" name="aaa" value="取消" onClick="javascript:window.close()"> 
      </td>
		</tr></form>
		</table>
	</td>
  </tr>
  </table>
  <%else%>
<table border="0" width="400" id="table2" bgcolor="#FF9900">
	<tr>
		<td>参数错误！</td>
	</tr>
</table>
<%End if%>
<table border="0" width="400" id="table2" bgcolor="#FF9900">
	<tr>
		<td>　</td>
	</tr>
</table>