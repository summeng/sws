<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<script language="JavaScript">
function closewindow()
{ 
  window.close();
}
</script>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
    <%Call OpenConn 
    dim id,gbook1
    id=trim(request("id"))
    set rs=server.createobject("adodb.recordset")
    sql="select * from [DNJY_gbook] where id="&cstr(id)
    rs.open sql,conn,1,1
    if rs.eof and rs.bof then
	response.write "没有留言"
	response.end
   	end if

   	%>
   	<title>留言与回复</title>
    <style type="text/css">
<!--
.style1 {color: #FFFFFF}
-->
    </style>
<body topmargin="3">

<div align="center">
  <center>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#CCCCCC" width="400" height="162">
  <tr>
    <td width="519" height="162">
    <div align="center">
      <center>
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
      <tr>
        <td height="20" background="../img/greenback.gif"> <p align="left" class="style1"><span style="font-size: 9pt; font-weight: 700">&nbsp;客户<font color="#FF0000"><%=rs("username")%></font>于<%=rs("fbsj")%>的留言</span></td>
      </tr>
      <tr>
        <td height="25"><div ><b>问：</b><%=rs("gbook1")%></div></td>
        </tr>
      <tr>
        <td><table width="90%"  border="0"  cellpadding="0" cellspacing="0">
          <tr>
            <td><p><font color="#FF0000"><b>站长回复:</b><%If rs("gbook2")<>"" then%><%=rs("gbook2")%><br>（回复时间：<%=rs("hfsj")%>）<%else%>抱歉，站长尚未回复，请过后关注。<%End if%></font></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="30"><div align="right"><%=title%>→<br>
          <%=web%>
        </div></td>
      </tr>
      <tr>
        <td height="1" bgcolor="#E3E3E3"></td>
      </tr>
      <tr>
        <td height="30">
        <p align="right">
        <b>
        <font style="FONT-SIZE: 12px; COLOR: #000066; FONT-FAMILY: 宋体">
        <span style="FONT-SIZE: 9pt; FONT-FAMILY: 宋体">
        <a class="yellow" href="javascript:closewindow();" target="_self" style="text-decoration: none">
        关闭窗口</a></span></font></b></td>
      </tr>
    </table>
      </center>
    </div>
    </td>
  </tr>
</table>
  </center>
</div>
<%
Call CloseRs(rs)
Call CloseDB(conn)
%>
