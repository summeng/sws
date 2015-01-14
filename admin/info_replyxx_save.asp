<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../Include/err.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
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
<%dim sdays,hfje,urlpage,hfxxzg
hfxxzg=request("hfxxzg")
If hfxxzg=1 Then
Call OpenConn
conn.execute "UPDATE DNJY_info SET hfje=0 WHERE  hfxx=1 and "&DN_Now&">=dqsj" '同时向用户更新帐户金额
conn.execute "UPDATE DNJY_info SET hfxx=0 WHERE  hfxx=1 and "&DN_Now&">=dqsj" '同时向用户更新帐户金额
End If
Call CloseDB(conn)
response.write "<p align=""center"">操作成功！</p>"
response.end
%>
