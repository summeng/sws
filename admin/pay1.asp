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
<%call checkmanage("02")%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<style type="text/css">
/*提示信息*/
.info {position:relative;background:#;color:#666; text-decoration:none;font-size:12px;width:0px;height:0px;text-align:center;border:0px solid #ccc;line-height:18px;}/*设置链接的属性,一定要设置为relative才能使提示层跟着链接走*/
.info:hover {background:#eee;color:#333;}
.info span {display: none }/*设置正常下的span为隐藏状态*/
.info:hover span /*设置hover下的span属性为呈现状态,并设置提示层的位置*/{display:block;position:absolute;top:0px;left:25px;width:200px;
border:1px solid #DC5826; background:#FFF3B4; color:#333;padding:5px;text-align:left;}
</style>
</head>
<BODY>

<%Call OpenConn
dim action,show,er
action=request("ok")
if action="" then
Set rs = conn.Execute("select * from DNJY_pay") 
%>
<form action=pay1.asp method=post name=pay>
<table width="98%" border="1" style="border-collapse: collapse; border-style: dotted; border-width: 0px" bordercolor="#333333" cellspacing="0" cellpadding="2">
<tr><td height=35  colspan=2> <a href=http://www.alipay.com target="_blank">支付宝</a> </td></tr>
<tr><td height=35  width=20% align=right >支付宝帐户&nbsp;</td><td  > <input type="password" value="<%=rs("alipay_id")%>" name=alipay_id size="30"> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>填写您的支付宝帐户帐号类型应与到帐方式对应否则支付过程出错！</span></a>&nbsp;&nbsp;<a href=http://www.alipay.com target="_blank">申请</a></td></tr>
<tr><td height=35  width=20% align=right >安全校验码 &nbsp;</td><td  > <input type="password" value="<%=rs("alipay_securityCode")%>" name=alipay_securityCode size="30"> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>请登陆支付宝网站，在“商家服务”中查得安全校验码</span></a></td></tr>
<tr><td height=35  width=20% align=right >合作伙伴ID &nbsp;</td><td  > <input type="password" value="<%=rs("partner")%>" name=partner size="30"> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>请登陆支付宝网站，在“商家服务”中查得帐户合作者身份（partnerID）</span></a></td></tr>
<tr><td height=35  width=20% align=right >到帐方式 &nbsp;</td><td  >
<%if rs("alipay_dz")=0 then%>
<input type="radio" name="alipay_dz" value="0" id="alipay_dz" checked>
即时到帐收款&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="alipay_dz" value="1" id="alipay_dz">
担保交易收款&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="alipay_dz" value="2" id="alipay_dz">
双功能收款&nbsp;&nbsp;&nbsp;&nbsp;
<%ElseIf rs("alipay_dz")=1 then%>                         
<input type="radio" name="alipay_dz" value="0" id="alipay_dz">
即时到帐收款&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="alipay_dz" value="1" id="alipay_dz" checked>
担保交易收款&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="alipay_dz" value="2" id="alipay_dz">
双功能收款&nbsp;&nbsp;&nbsp;&nbsp;
<%else%>
<input type="radio" name="alipay_dz" value="0" id="alipay_dz">
即时到帐收款
<input type="radio" name="alipay_dz" value="1" id="alipay_dz">
担保交易收款
<input type="radio" name="alipay_dz" value="2" id="alipay_dz" checked>
双功能收款
<%end if%>  &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>收款到帐方式必须与支付宝帐户申请的商家服务类型相应</span></a></td></tr>
<tr><td height=35  width=20% align=right >启用开关 &nbsp;</td><td  >
<%if rs("alipay_kg")=0 then%>
<input type="radio" name="alipay_kg" value="0" id="alipay_kg" checked>
关闭&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="alipay_kg" value="1" id="alipay_kg">
启用
<%else%>                         
<input type="radio" name="alipay_kg" value="0" id="alipay_kg">
关闭&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="alipay_kg" value="1" id="alipay_kg" checked>
启用<%end if%>  &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>启用或关闭此收款方式</span></a></td></tr>


<tr><td height=35  colspan=2> <a href=http://www.tenpay.com target="_blank">财付通</a> </td></tr>
<tr><td height=35  width=20% align=right >财付通商户号&nbsp;</td><td  > <input type="password" value="<%=rs("tenpay_id")%>" name=tenpay_id size="30"> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>填写您的财付通帐户</span></a>&nbsp;&nbsp;<a href=http://www.tenpay.com target="_blank">申请</a></td></tr>
<tr><td height=35  width=20% align=right >财付通商户密钥 &nbsp;</td><td  > <input type="password" value="<%=rs("tenpay_key")%>" name=tenpay_key size="30"> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>应与财付通网站用户登录后台设置的密钥相同</span></a></td></tr>
<tr><td height=35  width=20% align=right >订单服务说明 &nbsp;</td><td  > <input type=text value="<%=rs("tenpay_name")%>" name=tenpay_name size="30"> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>填写商品或者服务的说明，此项必须填写</span></a></td></tr>
<tr><td height=35  width=20% align=right >到帐方式 &nbsp;</td><td  >
<%if rs("tenpay_dz")=0 then%>
<input type="radio" name="tenpay_dz" value="0" id="tenpay_dz" checked>
即时到帐&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="tenpay_dz" value="1" id="tenpay_dz">
担保交易
<%else%>                         
<input type="radio" name="tenpay_dz" value="0" id="tenpay_dz">
即时到帐&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="tenpay_dz" value="1" id="tenpay_dz" checked>
担保交易<%end if%>  &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>收款到帐方式必须与财付通帐户申请的类型相应</span></a></td></tr>
<tr><td height=35  width=20% align=right >启用开关 &nbsp;</td><td  >
<%if rs("tenpay_kg")=0 then%>
<input type="radio" name="tenpay_kg" value="0" id="tenpay_kg" checked>
关闭&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="tenpay_kg" value="1" id="tenpay_kg">
启用
<%else%>                         
<input type="radio" name="tenpay_kg" value="0" id="tenpay_kg">
关闭&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="tenpay_kg" value="1" id="tenpay_kg" checked>
启用<%end if%>  &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>启用或关闭此收款方式</span></a></td></tr>
<tr><td height=35  width=20% align=right >公测帐号 &nbsp;</td><td  > 注：财付通提供的即时到帐、中介担保及双接口通用测试帐号<br>
商户号：1900000113<br>
商户名称：自助商户测试帐户<br>
密钥：e82573dc7e6136ba414f2e2affbe39fa<br>
注意：请不要支付大额金额到以上测试帐号（腾讯财付通恕不退款）,切记：正式上线时用属于您的商户号及密钥替换。&nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>财付通提供的即时到帐、中介担保及双接口通用测试帐号</span></a></td></tr>

<tr><td height=35  colspan=2> <a href=http://www.yeepay.com target="_blank">易宝支付</a>&nbsp;</td></tr> 
<tr><td height=35  width=20% align=right>商户编号&nbsp;</td><td  > <input type="password" value="<%=rs("yeepayid")%>" name=yeepayid size="30"> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>填写您在易宝支付申请的商户ID</span></a>&nbsp;&nbsp;<a href=http://www.yeepay.com target="_blank">申请</a></td></tr>
<tr><td height=35  width=20% align=right >密钥(key)&nbsp;</td><td  > <input type="password" value="<%=rs("yeepaykey")%>" name=yeepaykey size="60"> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>应与易宝支付的密钥一致</span></a></td></tr>

<tr><td height=35  width=20% align=right >启用开关 &nbsp;</td><td  >
<%if rs("yeepay_kg")=0 then%>
<input type="radio" name="yeepay_kg" value="0" id="yeepay_kg" checked>
关闭&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="yeepay_kg" value="1" id="yeepay_kg">
启用
<%else%>                         
<input type="radio" name="yeepay_kg" value="0" id="yeepay_kg">
关闭&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="yeepay_kg" value="1" id="yeepay_kg" checked>
启用<%end if%>  &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>启用或关闭此收款方式</span></a></td></tr>


<tr><td height=35  colspan=2 > <a href=http://www.chinabank.com.cn target="_blank">网银在线</a>&nbsp;</td></tr>
<tr><td height=35  width=20% align=right >商 户 号&nbsp;</td><td  > <input type="password" value="<%=rs("chinaid")%>" name=chinaid size="30"> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>填写您在网银在线申请的商户号一般为4或5位数字</span></a>&nbsp;&nbsp;<a href=http://www.chinabank.com.cn target="_blank">申请</a></td></tr>
<tr><td height=35  width=20% align=right >MD5 私钥&nbsp;</td><td  > <input type="password" value="<%=rs("chinakey")%>" name=chinakey size="30"> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>应与网银后台设置的密钥一致</span></a></td></tr>
<tr><td height=35  width=20% align=right rowspan=1 >返回地址&nbsp;</td><td  > <input type=text value="http://<%=web%>/<%=strInstallDir%>chinabank/Receive.asp" name=chinaback size=45> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>在线支付成功后的返回通知页面系统已设置此处没意义</span></a>&nbsp;</td></tr>
<tr><td height=35  width=20% align=right >启用开关 &nbsp;</td><td  >
<%if rs("china_kg")=0 then%>
<input type="radio" name="china_kg" value="0" id="china_kg" checked>
关闭&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="china_kg" value="1" id="china_kg">
启用
<%else%>                         
<input type="radio" name="china_kg" value="0" id="china_kg">
关闭&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="china_kg" value="1" id="china_kg" checked>
启用<%end if%>  &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>启用或关闭此收款方式</span></a></td></tr>




<tr><td height=35  colspan=2><INPUT name="ok" TYPE="hidden" value="ok"><INPUT name=action TYPE="submit" value="保存设置">&nbsp;</td></tr>
</table>
</form>
<%
rs.close
conn.close
set rs=nothing
set conn=nothing
end if
              
if action="ok" then

if request.form("yeepay_kg")=1 and request.form("yeepayid")="" then er=1
if request.form("china_kg")=1 then
if request.form("chinaid")="" or request.form("chinakey")="" then er=2
end if
if request.form("alipay_kg")=1 then
if request.form("alipay_id")="" or request.form("alipay_securityCode")="" or request.form("partner")="" then er=7
end if
if request.form("tenpay_kg")=1 then
if request.form("tenpay_id")="" or request.form("tenpay_key")="" or request.form("tenpay_name")="" then er=9
end if
if er=1 then
response.write "<script language='javascript'>"
response.write "alert('出错了，您选择了“易宝支付”，必须填写yeepaypay商户编号！');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
if er=2 then
response.write "<script language='javascript'>"
response.write "alert('出错了，您选择了“网银在线”，必须填写网银商户ID和MD5私钥！');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

if er=7 then
response.write "<script language='javascript'>"
response.write "alert('出错了，您选择了“支付宝”，必须填写支付宝帐户、安全校验码、合作伙伴ID！');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

if er=9 then
response.write "<script language='javascript'>"
response.write "alert('出错了，您选择了“财付通”，必须填写财付通商户帐号、密钥、产品和服务说明！');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from DNJY_pay"
rs.open sql,conn,1,3

rs("yeepayid")=request.form("yeepayid")
rs("yeepaykey")=request.form("yeepaykey")
'rs("yeepay_dz")=request.form("yeepay_dz")
rs("yeepay_kg")=request.form("yeepay_kg")

rs("chinaid")=request.form("chinaid")
rs("chinakey")=request.form("chinakey")
rs("chinaback")=request.form("chinaback")
'rs("china_dz")=request.form("china_dz")
rs("china_kg")=request.form("china_kg")

rs("alipay_id")=Trim(request.form("alipay_id"))
rs("alipay_securityCode")=Trim(request.form("alipay_securityCode"))
rs("partner")=Trim(request.form("partner"))
rs("alipay_dz")=request.form("alipay_dz")
rs("alipay_kg")=request.form("alipay_kg")

rs("tenpay_id")=request.form("tenpay_id")
rs("tenpay_name")=request.form("tenpay_name")
rs("tenpay_key")=request.form("tenpay_key")
rs("tenpay_dz")=request.form("tenpay_dz")
rs("tenpay_kg")=request.form("tenpay_kg")

rs.update
Call CloseRs(rs)
Call CloseDB(conn)
response.write "<script language='javascript'>"
response.write "alert('操作成功，您设置的在线支付信息已保存！');"
response.write "location.href='pay1.asp';"
response.write "</script>"
end if%>
