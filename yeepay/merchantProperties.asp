<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<%
'**************************************************
'* @Description �ױ�֧����Ʒͨ��֧���ӿڷ���  
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
	
	'�̻����p1_MerId,�Լ���ԿmerchantKey ��Ҫ���ױ�֧��ƽ̨���
	p1_MerId		= ""&yeepayid&"" '��ʽʹ��
	'p1_MerId		= "10001126856"	'����ʹ��
	merchantKey	= ""&yeepaykey&"" '��ʽʹ��
	'merchantKey	= "69cl522AV6q613Ii4W6u8K6XuW8vM1N6bFgyv769220IuYe9u37N4y7rI4Pl" '����ʹ��
	logName = "YeePay_HTML"
%>