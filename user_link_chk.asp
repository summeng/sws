<!--#include file="conn.asp"-->
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
<%Call OpenConn
dim id,i,str2,web,url,logo,webabout,strInstallDir
id=trim(request("selectedid"))
web=trim(request("web"))
url=trim(request("url"))
webabout=trim(request("webabout"))
id=trim(request("selectedid"))
str2=split(id,",")
for i=0 to ubound(str2)
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_comlink] where id="&trim(str2(i))
rs.open sql,conn,1,3
rs("web")=web
rs("url")=url
rs("webabout")=webabout
rs("jrsj")=now()
rs.update
next
response.write "修改成功！！"
response.write "<meta http-equiv=refresh content=""1;URL=user_link_e.asp?selectedid="&id&""">"
Call CloseRs(rs)
Call CloseDB(conn)%>