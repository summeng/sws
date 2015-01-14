<!--#include file="class/alipay_notify.asp"-->
<html>
<head>
	<META http-equiv=Content-Type content="text/html; charset=gb2312">
<%
' 功能：支付宝页面跳转同步通知页面
' 版本：3.3
' 日期：2012-07-17
' 说明：
' 以下代码只是为了方便商户测试而提供的样例代码，商户可以根据自己网站的需要，按照技术文档编写,并非一定要使用该代码。
' 该代码仅供学习和研究支付宝接口使用，只是提供一个参考。
	
' //////////////页面功能说明//////////////
' 该页面可在本机电脑测试
' 可放入HTML等美化页面的代码、商户业务逻辑程序代码
' 该页面可以使用ASP开发工具调试，也可以使用写文本函数LogResult进行调试，该函数已被默认关闭，见alipay_notify.asp中的函数VerifyReturn
' （该接口的注意事项，如果没有那么该行删除）。
'////////////////////////////////////////
%>
<%
'计算得出通知验证结果
Set objNotify = New AlipayNotify
sVerifyResult = objNotify.VerifyReturn()

If sVerifyResult Then	'验证成功
	'*********************************************************************
	'请在这里加上商户的业务逻辑程序代码
	
	'——请根据您的业务逻辑来编写程序（以下代码仅作参考）——
    '获取支付宝的通知返回参数，可参考技术文档中页面跳转同步通知参数列表

	'商户订单号
	out_trade_no = Request.QueryString("out_trade_no")

	'支付宝交易号
	trade_no = Request.QueryString("trade_no")

	'交易状态
	trade_status = Request.QueryString("trade_status")

	
	If Request.QueryString("trade_status") = "WAIT_SELLER_SEND_GOODS" Then
	'判断是否在商户网站中已经做过了这次通知返回的处理
		'如果没有做过处理，那么执行商户的业务程序
		'如果有做过处理，那么不执行商户的业务程序
	ElseIf request.QueryString("trade_status") = "TRADE_FINISHED" Then
		'判断该笔订单是否在商户网站中已经做过处理
			'如果没有做过处理，根据订单号（out_trade_no）在商户网站的订单系统中查到该笔订单的详细，并执行商户的业务程序
			'如果有做过处理，不执行商户的业务程序
	Else
		Response.Write "trade_status="&Request.QueryString("trade_status")
	End If

	Response.Write "<center>验证成功。付款操作完成。</center><br>"

	'——请根据您的业务逻辑来编写程序（以上代码仅作参考）——
'回写数据库支付状态=========================================
if request.QueryString("trade_status") = "WAIT_SELLER_SEND_GOODS" then
order_no= request.QueryString("out_trade_no")	'获取订单号，支付宝系统产生（或如果在alipayto.asp中设置用商家网站产生的）
total_fee= request.QueryString("price")		'获取总金额

Dim username,aliname
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_order] where ddhm='"&order_no&"' and cash="&total_fee&"" 
rs.open sql,conn,1,3
username=rs("username")
aliname=rs("aliname")
rs("zfqr")=1'同时更新用户订单付款状态
rs("zffs")=1'同时更新用户订单付款方式
rs("admincl")=1'同时更新用户订单审核状态
rs("ddshsj")=now()
rs("Payment")="支付宝"
rs.update
Call CloseRs(rs)

set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_Financial] "
rs.open sql,conn,1,3
rs.addnew
rs("username")=username
rs("aliname")=aliname
rs("ywje")=total_fee
rs("ywlx")="入帐"
rs("czbz")=""&order_no&" 订单金额入帐"
rs("czz")="在线支付"
rs("admincl")=1
rs("addTime")=now()
rs.update
Call CloseRs(rs)

conn.execute "UPDATE DNJY_user SET Amount=Amount+"&total_fee&" WHERE username='"&username&"'" '同时向用户更新帐户金额
Call CloseDB(conn)
End if
'回写数据库支付状态end===================================	
	'*********************************************************************
else '验证失败
    '如要调试，请看alipay_notify.asp页面的VerifyReturn函数，比对sign和mysign的值是否相等，或者检查responseTxt有没有返回true
    response.Write "<center>验证失败</center>"
end if
%>
<title>支付宝标准双接口</title>
<style type="text/css">
            .font_content{
                font-family:"宋体";
                font-size:14px;
                color:#FF6600;
            }
            .font_title{
                font-family:"宋体";
                font-size:16px;
                color:#FF0000;
                font-weight:bold;
            }
            table{
                border: 1px solid #CCCCCC;
            }
        </style>
</head>
<body>
<table align="center" width="700" cellpadding="5" cellspacing="0">
  <tr>
    <td align="center" class="font_title" colspan="2">通知返回</td>
  </tr>
  <tr>
    <td class="font_content" align="right">支付宝交易号：</td>
    <td class="font_content" align="left"><%=request.QueryString("trade_no")%></td>
  </tr>
  <tr>
    <td class="font_content" align="right">订单号：</td>
    <td class="font_content" align="left"><%=request.QueryString("out_trade_no")%></td>
  </tr>
  <tr>
    <td class="font_content" align="right">付款总金额：</td>
    <td class="font_content" align="left"><%=request.QueryString("total_fee")%></td>
  </tr>
  <tr>
    <td class="font_content" align="right">商品标题：</td>
    <td class="font_content" align="left"><%=request.QueryString("subject")%></td>
  </tr>
  <tr>
    <td class="font_content" align="right">商品描述：</td>
    <td class="font_content" align="left"><%=request.QueryString("body")%></td>
  </tr>
  <tr>
    <td class="font_content" align="right">买家账号：</td>
    <td class="font_content" align="left"><%=request.QueryString("buyer_email")%></td>
  </tr>
  <tr>
    <td class="font_content" align="right">交易状态：</td>
    <td class="font_content" align="left"><%if request.QueryString("trade_status") = "WAIT_SELLER_SEND_GOODS" or request.QueryString("trade_status") = "TRADE_FINISHED" Then
	response.Write "支付成功"
	else
    response.Write request.QueryString("trade_status")
	End if%></td>
  </tr>
</table>
</body>
</html>
