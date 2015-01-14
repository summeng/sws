<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=CONfig.ASP-->
<!--#include file="Include/Function.asp"-->
<!--#include file="Include/label.asp"-->
<!--#include file=usercookies.asp-->
<!--#include file="Include/tt_auto_lock.asp" -->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
dim rs2,username,aliname,ywje,ywlx,czbz,vipys,Amount,dqts,vip
username=trim(request.Form("username"))
if username="" then
response.write "<li>参数错误！"
response.end
end If
Call OpenConn
set rs2=server.createobject("adodb.recordset")
rs2.open"select username from DNJY_user where  username='"&username&"'",conn,1,1
if rs2.eof then
Response.Write ("<script language=javascript>alert('此用户不存在!');history.go(-1);</script>")
Call CloseRs(rs2)
Call CloseDB(conn)
response.end
end If
Call CloseRs(rs2)

username=trim(request.Form("username"))
aliname=trim(request("aliname"))
ywje=CInt(request.Form("ywje"))
vipys=CInt(request.Form("vipys"))
vipje=CInt(request.Form("vipje"))
ywlx=trim(request.Form("ywlx"))
czbz=HTMLEncode(trim(request.Form("czbz")))
Amount=CInt(request.Form("Amount"))
vip=CInt(request.Form("vip"))

if vip=1 then
Response.Write ("<script language=javascript>alert('您已是VIP会员，不用升级!');history.go(-1);</script>")
response.end
end If

if vipys<vipsj then
Response.Write ("<script language=javascript>alert('本站要求VIP资格至少一次性购买"&vipsj&"个月!');history.go(-1);</script>")
response.end
end If

if Amount<0 Or Amount<vipje*vipys then
Response.Write ("<script language=javascript>alert('您帐户的余额不足，请充值后再升级VIP会员资格!');history.go(-1);</script>")
response.end
end If

set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_Financial] "
rs.open sql,conn,1,3
rs.addnew
rs("username")=username
rs("aliname")=aliname
rs("ywje")=vipje*vipys
rs("ywlx")=ywlx
rs("czbz")=czbz
rs("czz")=username
rs("admincl")=1
rs("addTime")=now()
rs.update
Call CloseRs(rs)

conn.execute "UPDATE DNJY_user SET Amount=Amount-"&vipje*vipys&" WHERE username='"&username&"'" '同时向用户更新帐户金额
conn.execute "UPDATE DNJY_user SET tAmount=tAmount+"&vipje*vipys&" WHERE username='"&username&"'" '同时向用户更新消费总金额
dqts=vipys*30
dqts= dateadd("d", dqts, now)
conn.execute "UPDATE DNJY_user SET vip=1 WHERE username='"&username&"'" '同时升级为VIP资格
conn.execute "UPDATE DNJY_user SET vipdq='"&dqts&"' WHERE username='"&username&"'" '同时更新VIP资格到期时间
response.Write "<script language=javascript>alert('操作成功!');location.replace('upgradevip.asp');</script>"
Call CloseDB(conn)%>

