<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--  
 * ====================================================================
 *
 *                Send.asp 由网银在线技术支持提供
 *
 *  本页面接收来自上页所有订单信息,并提交支付订单到网线在线支付平台....
 *
 * 
 * ====================================================================
-->

<!--#include file="MD5.asp"-->
<%Dim v_mid,v_url,key,v_oid,curdate,ymd,hms,v_amount,v_moneytype,text,v_md5info,v_rcvname,v_rcvaddr,v_rcvtel,v_rcvpost,v_rcvemail,v_rcvmobile,v_ordername,v_orderaddr,v_ordertel,v_orderpost,v_orderemail,v_ordermobile,remark1,remark2'整合加上
dim chinaid,chinakey,chinaback
Call OpenConn
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_pay",conn,1,1
chinaid=rs("chinaid")
chinakey=rs("chinakey")
chinaback=rs("chinaback")    
Call CloseRs(rs)
Call CloseDB(conn)%>
<%
'****************************************	
	v_mid = chinaid					                 ' 商户号，测试商户号1001，密钥test。替换为自己的商户号(老版商户号为4位或5位,新版为8位)即可
    v_url = "http://"&web&"/"&strInstallDir&"chinabank/Receive.asp"    '商户自定义返回接收支付结果的页面 Receive.asp 为接收页面
	'v_url = chinaback                                '商户自定义返回接收支付结果的页面 Receive.asp 为接收页面
									
	key = chinakey									 ' 如果您还没有设置MD5密钥请登陆我们为您提供商户后台，地址：https://merchant3.chinabank.com.cn/
	'key = "test"												 ' 登陆后在上面的导航栏里可能找到“B2C”，在二级导航栏里有“MD5密钥设置” 
													 ' 建议您设置一个16位以上的密钥或更高，密钥最多64位，但设置16位已经足够了
'****************************************%>


<%
   if request("v_oid")<>"" then									'判断是否有传递订单号
   
		  v_oid = request("v_oid")
	  
	  else

		  curdate = now()										' 根据系统时间产生订单，格式：YYYYMMDD-v_mid-HMMSS
		  ymd = year(curdate)&month(curdate)&day(curdate)		' 年月日
		  hms = hour(curdate)&minute(curdate)&second(curdate)	' 分秒时

		  v_oid = ymd&"-"&v_mid&"-"&hms							' 推荐订单号构成格式为 年月日-商户号-小时分钟秒

	end if
	v_amount = request("v_amount")		' 订单金额
    v_amount = replace(v_amount,",","")
	v_moneytype = "CNY"					' 币种

	text = v_amount&v_moneytype&v_oid&v_mid&v_url&key	' 拼凑加密串

	v_md5info=Ucase(trim(md5(text)))					' 网银支付平台对MD5值只认大写字符串，所以小写的MD5值得转换为大写

'**********以下几项为可选信息,如果发送网银在线会保存此信息,使用和不使用都不影响支付！**************

	   v_rcvname = request("v_rcvname")			' 收货人
	   v_rcvaddr = request("v_rcvaddr")			' 收货地址
		v_rcvtel = request("v_rcvtel")			' 收货人电话
	   v_rcvpost = request("v_rcvpost")			' 收货人邮编
	  v_rcvemail = request("v_rcvemail")		' 收货人邮件
	 v_rcvmobile = request("v_rcvmobile")		' 收货人手机号

	 v_ordername = request("v_ordername")		' 订货人姓名
	 v_orderaddr = request("v_orderaddr")		' 订货人地址
	  v_ordertel = request("v_ordertel")		' 订货人电话
	 v_orderpost = request("v_orderpost")		' 订货人邮编
  	v_orderemail = request("v_orderemail")		' 订货人邮件
	v_ordermobile = request("v_ordermobile")	' 订货人手机号

		 remark1 = request("remark1")			' 备注字段1
		 remark2 = request("remark2")			' 备注字段2

%>

<!--以下信息为标准的 HTML 格式 + ASP 语言 拼凑而成的 网银在线 支付接口标准演示页面 无需修改-->

<html>

<body onLoad="javascript:document.E_FORM.submit()">
<form action="https://pay3.chinabank.com.cn/PayGate?encoding=utf-8" method="POST" name="E_FORM">


    
  <input type="hidden" name="v_md5info"    value="<%=v_md5info%>" size="100">
  <input type="hidden" name="v_mid"        value="<%=v_mid%>">
  <input type="hidden" name="v_oid"        value="<%=v_oid%>">
  <input type="hidden" name="v_amount"     value="<%=v_amount%>">
  <input type="hidden" name="v_moneytype"  value="<%=v_moneytype%>">
  <input type="hidden" name="v_url"        value="<%=v_url%>">
   
  <!--以下几项项为网上支付完成后，随支付反馈信息一同传给信息接收页 -->
    
  <input type="hidden"  name="remark1" value="<%=remark1%>">
  <input type="hidden"  name="remark2" value="<%=remark2%>">
    
<!--以下几项只是用来记录客户信息，可以不用，不影响支付 -->

	<input type="hidden"  name="v_rcvname"      value="<%=v_rcvname%>">
	<input type="hidden"  name="v_rcvaddr"      value="<%=v_rcvaddr%>">
	<input type="hidden"  name="v_rcvtel"       value="<%=v_rcvtel%>">
	<input type="hidden"  name="v_rcvpost"      value="<%=v_rcvpost%>">
	<input type="hidden"  name="v_rcvemail"     value="<%=v_rcvemail%>">
	<input type="hidden"  name="v_rcvmobile"    value="<%=v_rcvmobile%>">

	<input type="hidden"  name="v_ordername"    value="<%=v_ordername%>">
	<input type="hidden"  name="v_orderaddr"    value="<%=v_orderaddr%>">
	<input type="hidden"  name="v_ordertel"     value="<%=v_ordertel%>">
	<input type="hidden"  name="v_orderpost"    value="<%=v_orderpost%>">
	<input type="hidden"  name="v_orderemail"   value="<%=v_orderemail%>">
	<input type="hidden"  name="v_ordermobile"  value="<%=v_ordermobile%>">
  
  </form>

</body>
</html>