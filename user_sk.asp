<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=usercookies.asp-->
<!--#include file=Include/mail.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<style type="text/css">
<!--
.style2 {color: #333333}
.style3 {color: #FF0000}
-->
</style>
</head>

<body topmargin="0" leftmargin="0">
<!--#include file=top.asp-->
<div align="center">
  <center>
  <table width="980" height="371" border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" bgcolor="#FFFFFF" style="border-collapse: collapse">
    <tr>
      <td width="214" height="299" valign="top"><div align="center">
            <!--#include file=userleft.asp-->　
        </div></td>
		<td width="5">&nbsp;</td>
      <td width="760" height="299" align="center" valign="top">
           <table align="center">本站收款帐号，请留意下面注意事项<p>
            </table>
<%dim partner,alipay_id,chinaid,chinakey,chinaback,china_dz,china_kg,yeepayid,yeepaykey,yeepay_dz,yeepay_kg,pay,alipay_securityCode,aliorder,alipay_name,alipay_dz,alipay_kg,tenpay_id,tenpay_key,tenpay_name,tenpay_dz,tenpay_kg,ddhm,alimoney,hkuse,dianhua,CompPhone,Email,aliname,RecAddress,Zipcode
ddhm="1000000"
alimoney=0.01
hkuse="支付测试"
dianhua="1234567"
CompPhone="13000000000"
Email="mail@163.com"
aliname="您的尊姓大名"
RecAddress=""
Zipcode="100010"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_pay",conn,1,1   
pay= rs("pay")
chinaid=rs("chinaid")
chinakey=rs("chinakey")
chinaback=rs("chinaback")
china_dz=rs("china_dz")
china_kg=rs("china_kg")
alipay_id=rs("alipay_id")
alipay_securityCode=rs("alipay_securityCode")
partner=rs("partner")
alipay_dz=rs("alipay_dz")
alipay_kg=rs("alipay_kg")
yeepayid=rs("yeepayid")
yeepaykey=rs("yeepaykey")
yeepay_dz=rs("yeepay_dz")
yeepay_kg=rs("yeepay_kg")
tenpay_id=rs("tenpay_id")
tenpay_key=rs("tenpay_key")
tenpay_name=rs("tenpay_name")
tenpay_dz=rs("tenpay_dz")
tenpay_kg=rs("tenpay_kg")    
Call CloseRs(rs)

'在线支付开始
dim inBillNo,total
inBillNo=ddhm '订单号
total=alimoney '支付金额
%>
<table border="0" cellspacing="1" cellpadding="0" align="center" valign=absmiddle width=100%>
<tr><td  width=50% height=120>
<br>
<%if alipay_kg=1 then'=====支付宝支付开始%>                            
<table border="0" cellspacing="1" cellpadding="0" align="center" valign=absmiddle width=760>
<tr><td  width=20% height=80>
<%If alipay_dz=0 then%>
<form method="post" action="alipay/js_asp_gb/index.asp?ddhm=<%=ddhm%>&hkuse=<%=hkuse%>&alimoney=<%=alimoney%>" target=new><input type=submit  value="支付宝支付--即时到帐"></form>
<%ElseIf alipay_dz=1 then%>
<form method="post" action="alipay/db_asp_gb/index.asp?ddhm=<%=ddhm%>&hkuse=<%=hkuse%>&alimoney=<%=alimoney%>" target=new><input type=submit  value="支付宝支付--担保交易收款"></form>
<%else%>
<form method="post" action="alipay/sj_asp_gb/index.asp?ddhm=<%=ddhm%>&hkuse=<%=hkuse%>&alimoney=<%=alimoney%>" target=new><input type=submit  value="支付宝支付--双功能收款"></form>
<%End if%>
</td></tr>
</table>
<%End If

if tenpay_kg=1 then%>
<table border="0" cellspacing="1" cellpadding="0" align="center" valign=absmiddle width=760>
<tr><td  width=20% height=80>
<%If tenpay_dz=0 then%>
<form method="post" action="tenpay/b2c/index.asp?ddhm=<%=ddhm%>&hkuse=<%=hkuse%>&alimoney=<%=alimoney%>" target=new><input type=submit  value="财付通支付--即时到帐"></form>
<%else%>
<form method="post" action="tenpay/med/index.asp?ddhm=<%=ddhm%>&hkuse=<%=hkuse%>&alimoney=<%=alimoney%>" target=new><input type=submit  value="财付通支付--担保交易"></form>
<%End if%>
</td></tr>
</table>
<%End If
if china_kg=1 then
dim v_rcvtel,v_rcvmobile,v_orderemail,v_rcvname,v_rcvaddr,v_rcvpost,v_oid,v_amount,remark1
v_rcvtel=dianhua'电话
v_rcvmobile=CompPhone'手机
v_orderemail = email '信箱
v_rcvname =  aliname'收货人
v_rcvaddr =  RecAddress'地址
v_rcvpost = Zipcode'邮编
v_oid=inBillNo'订单号.
v_amount=replace(total,",","")'付款金额
remark1=webname&inBillNo&"号订单"'备注一
%>
<table border="0" cellspacing="1" cellpadding="0" align="center" valign=absmiddle width=760>
<tr><td  width=20% height=80>
<FORM name=formbill action="chinabank/Send.asp" method=post target=new>
<input type="submit" name="v_action" value="网银在线--即时到帐">
<input type="hidden" name="v_oid" value="<%=v_oid%>">
<input type="hidden" name="v_amount" value="<%=v_amount%>">
<input type="hidden" name="v_rcvtel" value="<%=v_rcvtel%>">
<input type="hidden" name="v_rcvmobile" value="<%=v_rcvmobile%>">
<input type="hidden" name="v_rcvname" value="<%=v_rcvname%>">
<input type="hidden" name="v_rcvaddr" value="<%=v_rcvaddr%>">
<input type="hidden" name="v_rcvpost" value="<%=v_rcvpost%>">
<input type="hidden" name="v_orderemail"  value="<%=v_orderemail%>">
<input type="hidden" name="remark1"  value="<%=remark1%>">
</form>
</td></tr>
</table>
<%End if

If yeepay_kg=1 then%>
<table border="0" cellspacing="1" cellpadding="0" align="center" valign=absmiddle width=100%>
<tr><td width=50% height=120>
<form method="post" action="yeepay/pay.asp?ddhm=<%=ddhm%>&hkuse=<%=hkuse%>&alimoney=<%=alimoney%>" target=new><input type=submit  value="易宝支付--即时到帐"></form>
</table>
</td></tr>
</table>
<%End if%>

<table border="0" cellspacing="1" cellpadding="0" align="center" valign=absmiddle width=100%>
<tr>如果无在线付款条件可选择如下传统汇款方式:<td  width=50%  height=150>
<center>
<%Dim er,bankN
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_pay ",conn,1,3
er=0
if rs("bank1")="" or isnull(rs("bank1")) then er=er+1
if rs("bank2")="" or isnull(rs("bank2")) then er=er+1
if rs("bank3")="" or isnull(rs("bank3")) then er=er+1
if rs("bank4")="" or isnull(rs("bank4")) then er=er+1
if rs("bank5")="" or isnull(rs("bank5")) then er=er+1
if rs("bank6")="" or isnull(rs("bank6")) then er=er+1
if rs("bank7")="" or isnull(rs("bank7")) then er=er+1
if rs("bank8")="" or isnull(rs("bank8")) then er=er+1
if rs("bank9")="" or isnull(rs("bank9")) then er=er+1
if rs("bank10")="" or isnull(rs("bank10")) then er=er+1
if rs("bank11")="" or isnull(rs("bank11")) then er=er+1
if rs("bank12")="" or isnull(rs("bank12")) then er=er+1
if rs("bank13")="" or isnull(rs("bank13")) then er=er+1
if rs("bank14")="" or isnull(rs("bank14")) then er=er+1
if rs("bank15")="" or isnull(rs("bank15")) then er=er+1


  if er=15 then '如果er=15，则表明没有全部银行帐号信息为空
response.write "<script language='javascript'>"
response.write "alert('您还没有添加银行帐号信息，请单击“确定”开始添加。！');"
response.write "location.href='pay2.asp?ok=edit';"
response.write "</script>"
response.end
  end if

'显示银行帐号信息
response.write "<table width=98% border='1'  style='border-collapse: collapse; border-style: dotted; border-width: 0px'  bordercolor='#333333' cellspacing='0' cellpadding='2'>"
response.write "<form action=pay2.asp method=post name=setup>"


for N=1 to 15
bankN="bank"&N
if rs(bankN)="" or isnull(rs(bankN)) then
else
response.write "<tr><td width=20% valign='middle' align='center'><img border=0 src='IMAGES/bank/"&bankN&".gif'></td><td  valign='middle'>"&rs(bankN)&"</td></tr>"
end if
next
response.write "</table>"
Call CloseRs(rs)
Call CloseDB(conn)%>
</td></tr>
</table>  

      </table></td>
      <td width="4" align="center" valign="top" bgcolor="#e6e6e6"></td>
    </tr>
    <tr>
      <td height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>