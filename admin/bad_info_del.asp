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
dim id,xid,i,u,str1,str2,tupian,objFSO,fileExt,sql1,rs1,xxid,bdback
bdback = Request.ServerVariables("HTTP_REFERER")
xid=trim(request("xid"))
id=trim(request("selectedid"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('没有选择记录！');" & "history.back()" & "</script>"
response.end
end if
str1=split(xid,",")
str2=split(id,",")
Call OpenConn
set rs1 = Server.CreateObject("ADODB.RecordSet")
if xid<>"" then
for i=0 to ubound(str1)
sql1="delete from [DNJY_Bad_info] where id="&id
rs1.open sql1,conn,1,3
next
else

for u=0 to ubound(str2)
sql1="delete from [DNJY_Bad_info] where id="&trim(str2(u))
rs1.open sql1,conn,1,3
next
end if

response.write "删除举报信息成功！"
response.redirect (bdback)
response.end
Call CloseRs(rs)
Call CloseRs(rs1)
Call CloseDB(conn)
%>
