<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/tt_auto_lock.asp" -->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%response.buffer="True"
dim user,username
dim problem,id
username=request.form("username")
problem=request.form("problem")
dim connuser,password
Call OpenConn
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_user] where username='"&UserName&"'"
rs.open sql,conn,1,1
if rs.eof then
Response.Write "没有此用户名！"
Call CloseRs(rs)
Call CloseDB(conn)
Response.End
End if
session("gpw_error")=session("gpw_error")+1
if rs.eof then response.write "<tr><td colspan=7 bgcolor=#ffffff>&middot;<font color=#ff0000>没有此会员名，无法找回密码！你已错误输入"&session("gpw_error")&"次，超限将锁定IP无法访问！</font></td></tr>"
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

</head>
<% if not rs.eof And rs("problem")<>"" then%>
<body topmargin="0" leftmargin="0">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="400" id="AutoNumber1" height="147">
    <tr> 
      <td width="500" height="25" bgcolor="#FF9900"> <p align="center">&nbsp; ..:::第二步：请输入你的答案:::...</td>
    </tr>
  <tr> 
    <td width="500" height="122"> 
      <table border="0" width="100%" id="table1" height="100">
             <form method="POST" name="form1" onSubmit="return form1_onsubmit()" action="RetakePassword_c.asp">
      <tbody>
        <tr>
          <td width="100%" align="center">　
             您的用户名：<%=rs("username")%> </td>
        </tr>
                <tr>
          <td width="100%" align="center">　
            密码问题：<%=rs("problem")%> </td>
        </tr>
       	<tr > 
         <td width="100%" align="center">　
         <span class="style1">请输入密码答案：</span><input class="inputa" type="answer1" maxlength="30" name="answer1" size="20" ></td>
                      </tr>      
        
      </tbody>
    </table>
	<input type="hidden" name="username" value="<%=rs("username")%>">
    <center><input type="submit" value=" 到下一步 " name="submit" >
      </td>
		</tr></form>
		</table>
	</td>
  </tr>
  </table>
<%else%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="380" id="AutoNumber1" height="70">
    <tr> 
      <td width="380" height="25" bgcolor="#FF9900"> <p align="center"> </td>
    </tr>
  <tr> 
    <td width="380" > 
   您的密码找回资料不全，无法找回密码，请与管理员联系！   
	</td>
  </tr>
  </table><%end If%>
<table border="0" width="380" id="table2" bgcolor="#FF9900">
	<tr>
		<td>　</td>
	</tr>
</table>
<%Call CloseRs(rs)
Call CloseDB(conn)%>