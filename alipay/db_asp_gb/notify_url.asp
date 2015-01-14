<!--#include file="alipay_config.asp"-->
<!--#include file="class/alipay_notify.asp"-->
<%
	'功能：支付宝主动通知调用的页面（服务器异步通知页面）
	'版本：3.1
	'日期：2010-11-23
	'说明：
	'以下代码只是为了方便商户测试而提供的样例代码，商户可以根据自己网站的需要，按照技术文档编写,并非一定要使用该代码。
	'该代码仅供学习和研究支付宝接口使用，只是提供一个参考。
	
''''''''''''页面功能说明'''''''''''''''''''
'创建该页面文件时，请留心该页面文件中无任何HTML代码及空格。
'该页面不能在本机电脑测试，请到服务器上做测试。请确保互联网上可以访问该页面。
'该页面调试工具请使用写文本函数log_result，该函数已被默认开启，见alipay_notify.asp中的函数notify_verify
'WAIT_BUYER_PAY(表示买家已在支付宝交易管理中产生了交易记录，但没有付款);
'WAIT_SELLER_SEND_GOODS(表示买家已在支付宝交易管理中产生了交易记录且付款成功，但卖家没有发货);
'WAIT_BUYER_CONFIRM_GOODS(表示卖家已经发了货，但买家还没有做确认收货的操作);
'TRADE_FINISHED(表示买家已经确认收货，这笔交易完成);
'该服务器异步通知页面面主要功能是：防止订单未更新。如果没有收到该页面打印的 success 信息，支付宝会在24小时内按一定的时间策略重发通知
'''''''''''''''''''''''''''''''''''''''''''
%>


<%Dim verify_result,order_no,total_fee
'计算得出通知验证结果
verify_result = notify_verify()

if verify_result then	'验证成功
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'请在这里加上商户的业务逻辑程序代码

'回写数据库支付状态=========================================
if request.Form("trade_status") = "WAIT_SELLER_SEND_GOODS" then
order_no= request.Form("out_trade_no")	'获取订单号，支付宝系统产生（或如果在alipayto.asp中设置用商家网站产生的）
total_fee= request.Form("price")		'获取总金额

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

	'——请根据您的业务逻辑来编写程序（以下代码仅作参考）——
    '获取支付宝的通知返回参数，可参考技术文档中服务器异步通知参数列表
    order_no		= request.Form("out_trade_no")	'获取订单号
    total_fee		= request.Form("price")			'获取总金额
	
	if request.Form("trade_status") = "WAIT_BUYER_PAY" then
	'该判断表示买家已在支付宝交易管理中产生了交易记录，但没有付款
		
		'判断该笔订单是否在商户网站中已经做过处理（可参考“集成教程”中“3.4返回数据处理”）
			'如果没有做过处理，根据订单号（out_trade_no）在商户网站的订单系统中查到该笔订单的详细，并执行商户的业务程序
			'如果有做过处理，不执行商户的业务程序
		
		response.Write "success"	'请不要修改或删除
		
		'调试用，写文本函数记录程序运行情况是否正常
        'log_result("这里写入想要调试的代码变量值，或其他运行的结果记录")
	elseif request.Form("trade_status") = "WAIT_SELLER_SEND_GOODS" then
	'该判断表示买家已在支付宝交易管理中产生了交易记录且付款成功，但卖家没有发货
		
		'判断该笔订单是否在商户网站中已经做过处理（可参考“集成教程”中“3.4返回数据处理”）
			'如果没有做过处理，根据订单号（out_trade_no）在商户网站的订单系统中查到该笔订单的详细，并执行商户的业务程序
			'如果有做过处理，不执行商户的业务程序
		
		response.Write "success"	'请不要修改或删除
		
		'调试用，写文本函数记录程序运行情况是否正常
        'log_result("这里写入想要调试的代码变量值，或其他运行的结果记录")
	elseif request.Form("trade_status") = "WAIT_BUYER_CONFIRM_GOODS" then
	'该判断表示卖家已经发了货，但买家还没有做确认收货的操作
	
		'判断该笔订单是否在商户网站中已经做过处理（可参考“集成教程”中“3.4返回数据处理”）
			'如果没有做过处理，根据订单号（out_trade_no）在商户网站的订单系统中查到该笔订单的详细，并执行商户的业务程序
			'如果有做过处理，不执行商户的业务程序
		
		response.Write "success"	'请不要修改或删除
		
		'调试用，写文本函数记录程序运行情况是否正常
        'log_result("这里写入想要调试的代码变量值，或其他运行的结果记录")
	elseif request.Form("trade_status") = "TRADE_FINISHED" then
	'该判断表示买家已经确认收货，这笔交易完成
	
		'判断该笔订单是否在商户网站中已经做过处理（可参考“集成教程”中“3.4返回数据处理”）
			'如果没有做过处理，根据订单号（out_trade_no）在商户网站的订单系统中查到该笔订单的详细，并执行商户的业务程序
			'如果有做过处理，不执行商户的业务程序
		
		response.Write "success"	'请不要修改或删除
		
		'调试用，写文本函数记录程序运行情况是否正常
        'log_result("这里写入想要调试的代码变量值，或其他运行的结果记录")
	else
		'其他状态判断。
		
		response.Write "success"	
		'调试用，写文本函数记录程序运行情况是否正常
		'log_result ("这里写入想要调试的代码变量值，或其他运行的结果记录")
	end if
	'——请根据您的业务逻辑来编写程序（以上代码仅作参考）——
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
else '验证失败
    response.Write "fail"
	'调试用，写文本函数记录程序运行情况是否正常
	'log_result ("这里写入想要调试的代码变量值，或其他运行的结果记录")
end if
%>