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
<%call checkmanage("03")
Call OpenConn
dim id,i,str2,username,objFSO,fileExt,sql1,rs1,sql2
username=request.cookies("dick")("username")
id=trim(request("selectedid"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('没有选择记录！');" &"window.location='admin_order.asp'" & "</script>"
response.end
end if
str2=split(id,",")
set rs=server.createobject("adodb.recordset")
for i=0 to ubound(str2)
sql="delete from [DNJY_order] where id="&cstr(str2(i))
rs.open sql,conn,1,3
next
response.write "<script language=JavaScript>" & chr(13) & "alert('删除成功！');" &"window.location='admin_order.asp'" & "</script>"
response.end
Call CloseRs(rs)
Call CloseDB(conn)
%>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
-->
</style>
