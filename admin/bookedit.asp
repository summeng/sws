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
<%Call OpenConn
dim id
id=trim(request("id"))
if not isnumeric(id) or id="" then
response.write "<li>参数错误！"
response.end
end if
set rs=server.createobject("adodb.recordset")
sql = "select gbook1 from DNJY_gbook where id="&cstr(id)
rs.open sql,conn,1,3
if request("chk")="1" then
savehf
end if
%>
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<title>修改留言</title>
<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="180" height="35">
    <form action="?id=<%=id%>&chk=1" method="POST">
    <tr>
      <td width="181" height="10" colspan="2">
      <p align="center">修改内容</td>
    </tr>
    <tr>
      <td width="181" height="20" colspan="2">
      <p align="center">
                    <textarea rows="18" name="memo" cols="35"><%=rs("gbook1")%></textarea></td>
    </tr>
    <tr>
      <td width="60" height="6"></td>
      <td width="121" height="6">
      <p align="center">
      <input type="submit" value="提交" name="B1" style="color: #FFFFFF; border: 1px solid #000000; background-color: #666666"></td>
    </tr>
    </form>
  </table>
  </center>
</div>
<%
sub savehf()
dim gbook1
gbook1=trim(request("memo"))
rs("gbook1")=gbook1
rs.update
Call CloseRs(rs)
response.write "<li>修改成功！"
cl
Call CloseDB(conn)
response.end
end sub

%>
<%sub cl()%>
<body onLoad="setTimeout(window.close, 2000)">
<%end sub%>
