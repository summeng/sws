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
dim id,xid,i,u,str1,str2,username,tupian,objFSO,fileExt,sql1,rs1,xxid,referer
referer = Request.ServerVariables("HTTP_REFERER")
username=request.cookies("dick")("username")
xid=trim(request("xid"))
id=trim(request("selectedid"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('没有选择记录！');" & "history.back()" & "</script>"
response.end
end if
str1=split(xid,",")
str2=split(id,",")

if xid<>"" then
for i=0 to ubound(str1)
Conn.Execute("Update DNJY_info Set hfcs=hfcs-1 where id="&cstr(xid))
set rs1 = Server.CreateObject("ADODB.RecordSet")
sql1="delete from [DNJY_reply] where id="&id
rs1.open sql1,conn,1,3
next
else

set rs1 = Server.CreateObject("ADODB.RecordSet")
for u=0 to ubound(str2)
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select xxid from [DNJY_reply] where id="&trim(str2(u))
rs.open sql,conn,1,1
if rs.eof then
exit for                                      
end if
xxid=rs("xxid")
Conn.Execute("Update DNJY_info Set hfcs=hfcs-1 where id="&xxid)
sql1="delete from [DNJY_reply] where id="&trim(str2(u))
rs1.open sql1,conn,1,3
next
end if

response.write "删除回复信息成功！"
If trim(request("back"))="xinxi" Then
response.write "<meta http-equiv=refresh content=""1;URL=infolist.asp"">"
else
response.redirect (referer)
End if

set rs=nothing
set rs1=nothing
Call CloseDB(conn)
response.end%>
