<%
'=====以下参数必须保持正确，由后台设置生成，改动需谨慎====www.ip126.com===

Dim API_QuickLogin,API_QQEnable,appid,appkey,callback,API_SinaEnable,SinaId,SinaKey,api_sinacallback,API_AlipayEnable,AlipayPartner,AlipayKey,api_alipayreturnurl
API_QuickLogin=false '是true否false开启自动创建帐号 
API_QQEnable=true '是否开启QQ登录
appid   = "1003024" 'QQ的APP ID
appkey  = "a55fdhj635453408d3kl43ef45932eb3839c4dc" 'QQ的KEY
callback = "http://test.ip126.com/api/qq/user.asp" 'QQ登录后跳转的地址
API_SinaEnable=true '是否开启新浪微博登录
SinaId="10w53598344h" '新浪微博登录App Key
SinaKey="36372etrth349il65eba1366534102acd5d5730" '新浪微博登录App Secret
api_sinacallback="http://test.ip126.com/api/sina/deal.asp" '新浪微博登录后跳转的地址
API_AlipayEnable=true '是否开启支付宝快捷登录
AlipayPartner="2038803454502242362" '支付宝合作者身份ID
AlipayKey="xgqghk8rgfghhgh5qfgt3epc3j502fjddsi5dnr" '安全检验码Key
api_alipayreturnurl="http://test.ip126.com/api/alipay/return_url.asp" '支付宝快捷登录后跳转的地址
%>
