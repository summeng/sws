<!--#include file="../../../conn.asp"-->
<!--#include file="../../../config.asp"-->
<%dim alipay_id,alipay_securityCode,partner,key,seller_email,notify_url,return_url,show_url,mainname,input_charset,sign_type,ddhm,hkuse,alimoney
Dim payment_type,out_trade_no,subject,price,quantity,logistics_fee,logistics_type,logistics_payment,body,receive_name,receive_zip,receive_phone,receive_mobile,sParaTemp,sHtml,receive_address,objSubmit,i,pos,nLen,itemName,itemValue,sPara,minmax,minmaxSlot,j,mark,temp,sParaSort,nCount,prestr,iPos,sItemName,sItemValue
Dim sVerifyResult,trade_no,trade_status,HTTPS_VERIFY_URL,objNotify,varItem,order_no,total_fee

Call OpenConn
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_pay",conn,1,1
alipay_id=rs("alipay_id")
alipay_securityCode=rs("alipay_securityCode")
partner=rs("partner")    
rs.close
set rs=Nothing
 %>
<%
' 配置文件
' 功能：设置帐户有关信息及返回路径
' 版本：3.3
' 日期：2012-07-13
' 说明：
' 以下代码只是为了方便商户测试而提供的样例代码，商户可以根据自己网站的需要，按照技术文档编写,并非一定要使用该代码。
' 该代码仅供学习和研究支付宝接口使用，只是提供一个参考。

' 提示：如何获取安全校验码和合作身份者ID
' 1.用您的签约支付宝账号登录支付宝网站(www.alipay.com)
' 2.点击“商家服务”(https://b.alipay.com/order/myOrder.htm)
' 3.点击“查询合作者身份(PID)”、“查询安全校验码(Key)”

' 安全校验码查看时，输入支付密码后，页面呈灰色的现象，怎么办？
' 解决方法：
' 1、检查浏览器配置，不让浏览器做弹框屏蔽设置
' 2、更换浏览器或电脑，重新登录查询。

'↓↓↓↓↓↓↓↓↓↓请在这里配置您的基本信息↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓

'合作身份者ID，以2088开头的16位纯数字
partner= ""&partner&""

'安全检验码，以数字和字母组成的32位字符
key= ""&alipay_securityCode&""


'↑↑↑↑↑↑↑↑↑↑请在这里配置您的基本信息↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑



'字符编码格式 目前支持 gbk 或 utf-8
input_charset= "gbk"

'签名方式 不需修改
sign_type= "MD5"
%>