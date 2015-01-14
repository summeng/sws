<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="config.asp"-->
<!--#include file="Include/Function.asp"-->
<!--#include file="Include/label.asp"-->
<!--#include file=usercookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<%Call OpenConn
dim id,d,user_d,i,username
i=1
id=trim(request("id"))
username=request.cookies("dick")("username")
set rs = Server.CreateObject("ADODB.RecordSet")
if request("dick")="chk" then
call edit_d()
response.end
else
sql="select d from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,1
if rs.eof then
response.write "参数错误！"
response.end
end if
user_d=rs("d")
if user_d=0 then
response.write "你还没有道具，请先购买道具！"
cl
response.end
end if
rs.close
%>
<body topmargin="0">
<title>信息修改-验证道具使用</title>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
<tr>
<td width="100%" height="25">
<p align="left">你现在还有<%=user_d%>个<font color="#0000FF">通过验证道具</font></td>
</tr>
<tr>
<form action="?id=<%=id%>&dick=chk" method="POST">
<td width="100%" height="25">
<p align="left">如果希望让该信息直接在页面显示请选择是否使用<font color="#0000FF">验证道具</font>：
<%if user_d>0 then%>&nbsp;&nbsp;&nbsp;
<select name="d">
<option selected value="0">不使用</option>
<option value="1">使用</option>
</select>&nbsp;&nbsp;                     
				  <input type="submit" value="提交" name="B1"><%else%>
<font color="#999999">没有道具</font>
<%end if%>
</td>
</form>
</tr>
</table>
<p align="left">此操作将减掉相应的道具数量使用一次需要一个道具</p>
<%
end if
sub edit_d()
sql="select d from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,3
d=rs("d")
if trim(request("d"))="0" then
response.write "你选择了不使用，请进行其他操作！"
cl
response.end
end if
if d=<0 then
response.write "没有足够的道具！"
cl
response.end
else
rs("d")=rs("d")-1
rs.update
end if
rs.close
sql="select yz,fbsj from [DNJY_info] where username='"&username&"' and id="&cstr(id)
rs.open sql,conn,1,3
rs("yz")=1
rs("fbsj")=now()
rs.update
response.write "使用道具成功，你的信息已经通过了验证！"
If Htmlhome=1 Then Call s_HomeHtml()'重新生成首页
rs.close
cl
response.end
end sub
%>
<%sub cl()%>
<body onLoad="setTimeout(window.close, 1000)">
<%end Sub
set rs=Nothing
Call CloseDB(conn)%>
