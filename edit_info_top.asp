<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="Include/Function.asp"-->
<!--#include file="Include/label.asp"-->
<!--#include file="SqlIn.Asp"-->
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
dim id,b,user_b,i,username
i=1
id=trim(request("id"))
username=request.cookies("dick")("username")
set rs = Server.CreateObject("ADODB.RecordSet")
if request("dick")="chk" then
call edit_b()
response.end
else
sql="select b from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,1
if rs.eof then
response.write "参数错误！"
response.end
end if
user_b=rs("b")
if user_b<1 then
response.write "你还没有道具，请先购买道具！"
cl
response.end
end if
rs.close
sql="select b from [DNJY_info] where username='"&username&"' and id="&cstr(id)
rs.open sql,conn,1,1
if rs.eof then
response.write "参数错误！"
response.end
end if
b=rs("b")
rs.close
%>
<body topmargin="0">
<title>信息修改-置顶道具使用</title>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
<tr>
<td width="100%" height="25">此前该信息<%if b>"0" then%>使用了<%=b%>个<%else%>没有使用<%end if%><font color="#0000FF">信息置顶道具
</font></td>
</tr>
<tr>
<td width="100%" height="25">你现在还有<%=user_b%>个<font color="#0000FF">信息置顶道具</font></td>
</tr>
<tr>

<form action="?dick=chk" name="form" method="POST">
<td width="100%" height="25">如果你想增加请选择数量
                  <%if user_b>0 then%>
                  <select name="b">
                  <option value="-<%=b%>">不使用</option>
				  <%for i=1 to user_b%>
				  <option value="<%=i%>"><%=i%></option>
				  <%next%>
				  </select>
				  <input type="hidden" name="id" value="<%=id%>">                     
				  <input class="inputb" type="submit" value="提交" name="B1">
                  <%else%><font color="#999999">没有道具</font>
                  <%end if%>
				  
</td>
</form>
</tr>
</table>
<p>此操作将减掉相应的道具数量</p>
<%
end if
sub edit_b()
sql="select b from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,3
b=rs("b")
if b<1 then
response.write "没有足够的道具！"
cl
response.end
else
rs("b")=rs("b")-int(trim(request.form("b")))
rs.update
end if
rs.close
sql="select b from [DNJY_info] where username='"&username&"' and id="&cstr(id)
rs.open sql,conn,1,3
if rs("b")=0 then
rs("b")=trim(request.form("b"))
else
rs("b")=rs("b")+int(trim(request.form("b")))
end if
rs.update
response.write "修改信息置顶道具成功！"
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
