<!--#include file="MD5.asp"-->

<%
'****************************************	' MD5密钥要跟订单提交页相同，如Send.asp里的 key = "test" ,修改""号内 test 为您的密钥
											' 如果您还没有设置MD5密钥请登陆我们为您提供商户后台，地址：https://merchant3.chinabank.com.cn/
	key = "test"							' 登陆后在上面的导航栏里可能找到“B2C”，在二级导航栏里有“MD5密钥设置” 
											' 建议您设置一个16位以上的密钥或更高，密钥最多64位，但设置16位已经足够了
'****************************************

v_oid=request("v_oid")'订单号
v_pmode=request("v_pmode")'银行名称 如：招商银行
v_pstatus=request("v_pstatus")'支付状态 如：20 支付成功，30 支付失败
v_pstring=request("v_pstring")'支付状态说明 如：支付成功
v_amount=request("v_amount")'支付金额
v_moneytype=request("v_moneytype")'币种 如：CNY

v_md5str=request("v_md5str")'MD5效验码

remark1=request("remark1")'备注1
remark2=request("remark2")'备注2


if v_md5str = "" then
	
	response.write("error")

	response.end '中断程序

end if

text = v_oid&v_pstatus&v_amount&v_moneytype&key '拼凑加密串

md5text = Ucase(trim(md5(text))) '生成MD5效验码

if md5text<>v_md5str then '与网银在线发送过来的MD5效验码对比，确保是网银在线发送的信息

	response.write("error") '告诉服务器验证失败，要求重发

	response.end '中断程序

else

	response.write("ok") '告诉服务器已经正确接收以及验证参数正确，要求停止发送

	if v_pstatus = "20" then

		'支付已经成功
		'此处加入商户系统的逻辑处理（例如判断金额，判断支付状态(20成功,30失败)，更新订单状态等等）......

	end if

end if
%>