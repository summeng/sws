<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="config.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Call OpenConn
set rs=server.createobject("adodb.recordset")
sql = "select b,fbsj from DNJY_info where b>0 order by b "&DN_OrderType&""
rs.open sql,conn,1,3
do while not rs.eof
dim aaa,bbb,ccc
aaa=rs("b")-datediff("d",rs("fbsj"),date())
if aaa>0 or aaa=0 then
rs("b")=aaa
rs("fbsj")=now()
rs.update
end if
ccc=ccc+1
if ccc>=100 then exit do
rs.movenext
loop
Call CloseRs(rs)
if b_y>0 then
Conn.Execute("Update DNJY_info Set b=0 where DateDiff("&DN_DatePart_D&",fbsj,"&DN_Now&")>="&b_y&" and b>0 and yz=1")
end if
if tui_y>0 then
Conn.Execute("Update DNJY_info Set tj=0 where DateDiff("&DN_DatePart_D&",fbsj,"&DN_Now&")>="&tui_y&" and tj>0 and yz=1")
end If
Call CloseDB(conn)%>