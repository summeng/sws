<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=../config.ASP-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
dim objFSO,fileExt,sql1,rs1,sql2,delpas,username
id=request("id")
username=request.cookies("dick")("username")
If request.cookies("dick")("username")="" then
response.Write "参数错误，请确认会员登录才能查询！"
response.end
End if
Call OpenConn
if request("key")="cs" And username<>"" then
%>
<form method="POST" action="?username=<%=username%>&cskey=ok&id=<%=id%>">
<p align="center">查询本信息发布人手机号码<br>将在您的帐户上扣除<%=jf_ck%>点积分</p>
<p align="center">目前您帐户有<%=conn.Execute("Select jf From DNJY_user Where username='"&username&"'")(0)%>点积分</p>
<p align="center"> <input type="submit" value="确定查询" name="B1"></p>
<p align="center"><a href="../help.asp?id=1&<%=C_ID%>" target=_blank>赚取积分方法</a></p>
<!--<p align="center"><a href="../help.asp?id=0&<%=C_ID%>" target=_blank>赚取积分方法</a></p><!--车网用-->
</form>
<%Elseif request("cskey")="ok" And conn.Execute("Select jf From DNJY_user Where username='"&username&"'")(0)>=jf_ck then
set rs=server.createobject("adodb.recordset")
sql="select CompPhone,name from [DNJY_info] where id="&cstr(id)
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then
conn.execute "UPDATE DNJY_user SET jf=jf-"&jf_ck&" WHERE username='"&username&"'" '同时向用户数据库减少积分
response.Write "<TABLE align=""center""><TR><TD><b>本信息联系人：</b>"&rs("name")&"<br><b>联系手机号码：</b>"&rs("CompPhone")&"<p>"
response.Write "目前您帐户余："&conn.Execute("Select jf From DNJY_user Where username='"&username&"'")(0)&"点积分</TD></TR></TABLE>"
response.Write "<p align=""center"">请勿刷新本页，否则重复扣分！</p>"
Call CloseRs(rs)
Call CloseDB(conn)
response.end
Else
response.Write "参数错误或您的积分不足，无法查询！"
end If
else
response.Write "参数错误无法查询，下列原因之一：1、您的积分不足。2、会员未登录。"
end If
Call CloseRs(rs)
Call CloseDB(conn)%>
