<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
dim id
id=trim(request("id"))
if not isnumeric(id) or id="" then
response.write "<li>参数错误！"
response.end
end If
Call OpenConn
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select a,b,c,d from [DNJY_user] where id="&cstr(id)
rs.open sql,conn,1,1
%>
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<title>修改道具数量</title>
<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#CCCCCC" width="180" height="35">
    <form action="user_editprops_save.asp?id=<%=id%>" method="POST">
    <tr>
      <td width="80" height="25">
      <p align="center"><font color="#FF0000">颜色道具：</font></td>
      <td width="101" height="25">
      <input type="text" name="a" size="10" value="<%=rs("a")%>" maxlength="10">
      <font color="#FF0000">个</font></td>
    </tr>
    <tr>
      <td width="80" height="25">
      <p align="center"><font color="#FF0000">置顶道具：</font></td>
      <td width="101" height="25">
      <input type="text" name="b" size="10" value="<%=rs("b")%>" maxlength="10">
      <font color="#FF0000">个</font></td>
    </tr>
    <tr>
      <td width="80" height="25">
      <p align="center"><font color="#FF0000">图片道具：</font></td>
      <td width="101" height="25">
      <input type="text" name="c" size="10" value="<%=rs("c")%>" maxlength="10">
      <font color="#FF0000">个</font></td>
    </tr>
    <tr>
      <td width="80" height="25">
      <p align="center"><font color="#FF0000">验证道具：</font></td>
      <td width="101" height="25">
      <input type="text" name="d" size="10" value="<%=rs("d")%>" maxlength="10">
      <font color="#FF0000">个</font></td>
    </tr>
    <tr>
      <td width="80" height="6"></td>
      <td width="101" height="6">
      <p align="center">
      <input type="submit" value="提交" name="B1" style="color: #FFFFFF; border: 1px solid #000000; background-color: #666666"></td>
    </tr>
    <tr>
      <td width="181" height="6" colspan="2">
      <p align="center"><font color="#FF0000">全部必须为数字</font></td>
    </tr>
    </form>
  </table>
  </center>
</div>
<%
Call CloseRs(rs)
Call CloseDB(conn)
%>
