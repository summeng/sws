<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--  
 * ====================================================================
 *
 *                 Receive.asp 由网银在线技术支持提供
 *
 *     本页面为支付完成后获取返回的参数及处理......
 *
 *
 * ====================================================================
-->
<!--#include file="MD5.asp"-->
<%dim chinaid,chinakey,chinaback
Call OpenConn
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from dnjy_pay",conn,1,1
chinaid=rs("chinaid")
chinakey=rs("chinakey")
chinaback=rs("chinaback")    
Call CloseRs(rs)
%>
<%Dim key,v_oid,v_pmode,v_pstatus,v_pstring,v_amount,v_moneytype,remark1,remark2,v_md5str,text,md5text'整合加上   
'****************************************	' MD5密钥要跟订单提交页相同，如Send.asp里的 key = "test" ,修改""号内 test 为您的密钥
											' 如果您还没有设置MD5密钥请登陆我们为您提供商户后台，地址：https://merchant3.chinabank.com.cn/
	key = chinakey							' 登陆后在上面的导航栏里可能找到“B2C”，在二级导航栏里有“MD5密钥设置”
											' 建议您设置一个16位以上的密钥或更高，密钥最多64位，但设置16位已经足够了
'****************************************

' 取得返回参数值
	v_oid=request("v_oid")                               ' 商户发送的v_oid定单编号
	v_pmode=request("v_pmode")                           ' 支付方式（字符串） 
	v_pstatus=request("v_pstatus")                       ' 支付状态 20（支付成功）;30（支付失败）
	v_pstring=request("v_pstring")                       ' 支付结果信息 支付完成（当v_pstatus=20时）；失败原因（当v_pstatus=30时）；
	v_amount=request("v_amount")                         ' 订单实际支付金额
	v_moneytype=request("v_moneytype")                   ' 订单实际支付币种
	remark1=request("remark1")                           ' 备注字段1
	remark2=request("remark2")                           ' 备注字段2
	v_md5str=request("v_md5str")                         ' 网银在线拼凑的Md5校验串


if request("v_md5str")="" then
	response.Write("v_md5str：空值")
	response.end
end if


'md5校验

	text = v_oid&v_pstatus&v_amount&v_moneytype&key

	md5text =Ucase(trim(md5(text)))    '商户拼凑的Md5校验串

	if md5text<>v_md5str then		' 网银在线拼凑的Md5校验串 与 商户拼凑的Md5校验串 进行对比
	'对比失败表示信息非网银在线返回的信息

	 response.write("校验失败,数据可疑")
      response.end
	else
	'对比成功表示信息是网银在线返回的信息
	 if v_pstatus=20 then

		'支付成功
		'此处加入商户系统的逻辑处理（例如判断金额，判断支付状态，更新订单状态等等）......
		'------------------------------

'回写数据库支付状态============================================================
Dim username,aliname
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_order] where ddhm='"&v_oid&"' and cash="&v_amount&"" 
rs.open sql,conn,1,3
username=rs("username")
aliname=rs("aliname")
rs("zfqr")=1'同时更新用户订单付款状态
rs("zffs")=1'同时更新用户订单付款方式
rs("admincl")=1'同时更新用户订单审核状态
rs("ddshsj")=now()
rs("Payment")="网银在线"
rs.update
Call CloseRs(rs)

set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_Financial] "
rs.open sql,conn,1,3
rs.addnew
rs("username")=username
rs("aliname")=aliname
rs("ywje")=v_amount
rs("ywlx")="入帐"
rs("czbz")=""&v_oid&" 订单金额入帐"
rs("czz")="在线支付"
rs("admincl")=1
rs("addTime")=now()
rs.update
Call CloseRs(rs)
conn.execute "UPDATE DNJY_user SET Amount=Amount+"&v_amount&" WHERE username='"&username&"'" '同时向用户更新帐户金额
Call CloseDB(conn)
'回写数据库支付状态end=========================================================

	end if
	end if%>

<!--
以下是打印出所有接收数据的结果，供编程人员参考
-->
<TABLE width=500 border=0 align="center" cellPadding=0 cellSpacing=0>
		  <TBODY>
			<TR> 
			  <TD vAlign=top align=middle> <div align="left"><B><FONT style="FONT-SIZE:14px">恭喜您，支付成功！</FONT></B></div></TD>
			</TR>
			<TR> 
			  <TD vAlign=top align=middle> <div align="left"><B><FONT style="FONT-SIZE:14px">MD5校验码:<%=v_md5str%></FONT></B></div></TD>
			</TR>
			<TR> 
			  <TD vAlign=top align=middle> <div align="left"><B><FONT style="FONT-SIZE: 14px">订单号:<%=v_oid%></FONT></B></div></TD>
			</TR>
			<TR> 
			  <TD vAlign=top align=middle> <div align="left"><B><FONT style="FONT-SIZE: 14px">支付卡种:<%=v_pmode%></FONT></B></div></TD>
			</TR>
			<TR> 
			  <TD vAlign=top align=middle> <div align="left"><B><FONT style="FONT-SIZE: 14px">支付结果:<%=v_pstring%></FONT></B></div></TD>
			</TR>
			<TR> 
			  <TD vAlign=top align=middle> <div align="left"><B><FONT style="FONT-SIZE: 14px">支付金额:<%=v_amount%></FONT></B></div></TD>
			</TR>
			<TR> 
			  <TD vAlign=top align=middle> <div align="left"><B><FONT style="FONT-SIZE: 14px">支付币种:<%=v_moneytype%></FONT></B></div></TD>
			</TR>
		  </TBODY>
		</TABLE>