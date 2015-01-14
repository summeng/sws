<!--#include file="./tenpay_config.asp"-->
<!--#include file="./classes/PayRequestHandler.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk">
<title>�Ƹ�ͨ˫�ӿ�֧������ʾ��</title>
</head>
<body>
<%
'---------------------------------------------------------
'�Ƹ�ͨ˫�ӿ�֧������ʾ�����̻����մ��ĵ����п�������
'---------------------------------------------------------
'��ȡ�ύ����Ʒ�۸�
order_price=trim(request("order_price"))

'��ȡ�ύ����Ʒ����
product_name=trim(request("product_name"))

'��ȡ�ύ�ı�ע��Ϣ
remarkexplain=trim(request("remarkexplain"))

'��ȡ�ύ�Ķ�����
out_trade_no=trim(request("order_no"))

'�ӿ�����
trade_mode=trim(request("trade_mode"))

'֧����ʽ
bank_type_value=trim(request("bank_type_value"))

'Dim total_fee
Dim desc

'��Ʒ�۸񣨰����˷ѣ����Է�Ϊ��λ
total_fee = csng(order_price*100)

'��Ʒ����
desc = "��Ʒ��" & product_name&",��ע��"&remarkexplain

'��������ʱ��
strNow = Now()
strNow = Year(strNow) & Right(("00" & Month(strNow)),2) & Right(("00" & Day(strNow)),2) & Right(("00" & Hour(strNow)),2) & Right(("00" &  Minute(strNow)),2) & Right(("00" & Second(strNow)),2)


'����֧���������
Dim reqHandler
Set reqHandler = new PayRequestHandler
reqHandler.setGateUrl("https://gw.tenpay.com/gateway/pay.htm")

'��ʼ��
reqHandler.init()

'������Կ
reqHandler.setKey(key)

'-----------------------------
'����֧������
'-----------------------------
reqHandler.setParameter "partner", partner		'�����̻���
reqHandler.setParameter "out_trade_no", out_trade_no				'�̻�������
reqHandler.setParameter "total_fee", total_fee				'��Ʒ�ܽ��,�Է�Ϊ��λ
reqHandler.setParameter "return_url", return_url			'�ص���ַ
reqHandler.setParameter "notify_url", notify_url			'֪ͨ��ַ
reqHandler.setParameter "body", desc	                    '��Ʒ����
reqHandler.setParameter "bank_type", bank_type_value						'��������
reqHandler.setParameter "fee_type", "1"						'���б���
reqHandler.setParameter "subject", desc             '��Ʒ����(�н齻��ʱ����)
reqHandler.setParameter "spbill_create_ip", Request.ServerVariables("REMOTE_ADDR")  '֧������IP

'ϵͳ��ѡ����
reqHandler.setParameter "sign_type", "MD5"        'ǩ����ʽ
reqHandler.setParameter "service_version", "1.0"  '�ӿڰ汾
reqHandler.setParameter "input_charset", "GBK"    '�ַ���
reqHandler.setParameter "sign_key_index", "1"     '��Կ���

'ҵ���ѡ����
reqHandler.setParameter "attach", ""                      '�������ݣ�ԭ������
reqHandler.setParameter "product_fee", ""                 '��Ʒ���ã����뱣֤transport_fee + product_fee=total_fee
reqHandler.setParameter "transport_fee", "0"            '�������ã����뱣֤transport_fee + product_fee=total_fee
reqHandler.setParameter "time_start", strNow            '��������ʱ�䣬��ʽΪyyyymmddhhmmss
reqHandler.setParameter "time_expire", ""                 '����ʧЧʱ�䣬��ʽΪyyyymmddhhmmss
reqHandler.setParameter "buyer_id", ""                    '�򷽲Ƹ�ͨ�˺�
reqHandler.setParameter "goods_tag", ""                   '��Ʒ���
reqHandler.setParameter "trade_mode", trade_mode                 '����ģʽ��1��ʱ����(Ĭ��)��2�н鵣����3��̨ѡ����ҽ�֧�������б�ѡ��
reqHandler.setParameter "transport_desc", ""              '����˵��
reqHandler.setParameter "trans_type", "1"                 '�������ͣ�1ʵ�ｻ�ף�2���⽻��
reqHandler.setParameter "agentid", ""                     'ƽ̨ID
reqHandler.setParameter "agent_type", ""                  '����ģʽ��0�޴���(Ĭ��)��1��ʾ������ģʽ��2��ʾ����ģʽ
reqHandler.setParameter "seller_id", ""                   '�����̻��ţ�Ϊ�����ͬ��partner

'�����URL
Dim reqUrl
reqUrl = reqHandler.getRequestURL()

'debug��Ϣ
Dim debugInfo
debugInfo = reqHandler.getDebugInfo()
'Response.Write("<br/>debugInfo:" & debugInfo & "<br/>")
'Response.Write("<br/>reqUrl:" & reqUrl & "<br/>")
'�ض��򵽲Ƹ�֧ͨ��
'reqHandler.doSend()

%>
<br/><TABLE align="center">
<TR>
	<TD><a href="<%=reqUrl%>"><img src="image/pay.jpg" width="212" height="55" border="0"></a><br>��ǰ�ĸ��ʽ���н鵣������</TD>
</TR>
</TABLE>
</body>
</html>