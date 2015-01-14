<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%dim server_v1,server_v2,errc
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))'本地路径
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))'服务器路径
if mid(server_v1,8,len(server_v2))<>server_v2 then
response.write ("<table cellSpacing=0 borderColorDark=#ffffff cellPadding=0  width=600 align=center borderColorLight=#0066CC border=1 bgcolor='#F0F0F0'>")
response.write  ("<tr align=center>")
response.write    ("<td height=36>你提交的路径有误，禁止从站点外部提交数据,请关闭窗口！</td>")
response.write  ("</tr>")
response.write ("<tr align=center>")
response.write ("<form><td height=30>")
response.write ("<INPUT TYPE='BUTTON' value='关闭窗口' onClick='window.close()'> ")
response.write ("</td></form>")
response.write ("</tr>")
response.write ("</table>")
response.end
end if
%>
<% 
if errc then 
response.write "<script language=""javascript"">" 
response.write "parent.alert('非法操作!禁止攻击站点服务器!');" 
response.write "self.location.href='Index.Asp';" 
response.write "</script>" 
response.end 
end if 
%>
