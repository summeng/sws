<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<%
'**************************************************
'* @Description 易宝支付产品通用支付接口范例  
'* @V3.0
'* @Author rui.xin
'**************************************************
dim yeepayid,yeepaykey
Call OpenConn
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_pay",conn,1,1
yeepayid=rs("yeepayid")
yeepaykey=rs("yeepaykey") 
Call CloseRs(rs)


	Dim p1_MerId
	Dim merchantKey,logName
	
	'商户编号p1_MerId,以及密钥merchantKey 需要从易宝支付平台获得
	p1_MerId		= ""&yeepayid&"" '正式使用
	'p1_MerId		= "10001126856"	'测试使用
	merchantKey	= ""&yeepaykey&"" '正式使用
	'merchantKey	= "69cl522AV6q613Ii4W6u8K6XuW8vM1N6bFgyv769220IuYe9u37N4y7rI4Pl" '测试使用
	logName = "YeePay_HTML"
%>