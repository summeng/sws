<!--#include file="./tenpay_config.asp"-->
<!--#include file="./classes/PayResponseHandler.asp"-->
<!--#include file="./util/tenpay_util.asp"-->
<%
'---------------------------------------------------------
'�Ƹ�ͨ˫�ӿڴ���ص�ʾ�����̻����մ�ʾ�����п�������
'---------------------------------------------------------

log_result("����ǰ̨֪ͨҳ�棡��������")


'����֧��Ӧ�����
Dim resHandler
Set resHandler = new PayResponseHandler
resHandler.setKey(key)

'�ж�ǩ��
If resHandler.isTenpaySign() = True Then

	Dim notify_id
	Dim transaction_id
	'Dim total_fee
	'Dim out_trade_no
	Dim discount
	Dim trade_state

	 '֪ͨid
	 notify_id = resHandler.getParameter("notify_id")

	'�̻����׵���
	out_trade_no = resHandler.getParameter("out_trade_no")	

	'�Ƹ�ͨ���׵���
	transaction_id = resHandler.getParameter("transaction_id")

	'��Ʒ���,�Է�Ϊ��λ
	total_fee = resHandler.getParameter("total_fee")
	
	'�����ʹ���ۿ�ȯ��discount��ֵ��total_fee+discount=ԭ�����total_fee
	discount = resHandler.getParameter("discount")
	
	'֧�����
	trade_state = resHandler.getParameter("trade_state")
	
	'����ģʽ��1��ʱ���ˣ�2�н鵣��
	trade_mode = resHandler.getParameter("trade_mode")
	'�ɻ�ȡ��������������
	'bank_type			��������,Ĭ�ϣ�BL
	'fee_type			�ֽ�֧������,Ŀǰֻ֧�������,Ĭ��ֵ��1-�����
	'input_charset		�ַ�����,ȡֵ��GBK��UTF-8��Ĭ�ϣ�GBK��
	'partner			�̻���,�ɲƸ�ͨͳһ�����10λ������(120XXXXXXX)��
	'product_fee		��Ʒ���ã���λ�֡������ֵ�����뱣֤transport_fee + product_fee=total_fee
	'sign_type			ǩ�����ͣ�ȡֵ��MD5��RSA��Ĭ�ϣ�MD5
	'time_end			֧�����ʱ��
	'transport_fee		�������ã���λ�֣�Ĭ��0�������ֵ�����뱣֤transport_fee +  product_fee = total_fee
	If "1" = trade_mode Then '��ʱ����
		If "0" = trade_state Then
	
	
	
	
			log_result("��ʱ����ǰ̨֪ͨ�ɹ�")
			Response.Write("success")
			Response.Write("<br>����"&out_trade_no&"��ʱ���ʸ���ɹ������ǻᾡ�촦����ҵ��лл������")'��ʱ����֧���ɹ�Success
	    '����ɹ�
		Else  
			log_result("��ʱ����ǰ̨֪ͨʧ��")
			Response.Write("��ʱ����֧��ʧ��trade_state=" & trade_state)'
		End If
	Else
		If "2"= trade_mode  Then    ' �н鵣�� 
			If  "0"= trade_state  Then
		
		
		
		
				log_result("�н鵣��ǰ̨֪ͨ�ɹ�")
				Response.Write("Success")'֧���ɹ�"trade_state=" + trade_state
				Response.Write("<br>����"&out_trade_no&"�н鵣������ɹ������ǻᾡ�촦����ҵ���뼰ʱȷ�ϸ��лл������")'֧���ɹ�Success"trade_state=" + trade_state
			Else 
				log_result("�н鵣��ǰ̨֪ͨʧ�ܣ�trade_state=" & trade_state)
				Response.Write("�н鵣��֧��ʧ��trade_state=" & trade_state)

	    	End if
	    End If
    End If
Else'ǩ��ʧ��
	Response.Write("ǩ��ǩ֤ʧ��")
	Dim debugInfo
	debugInfo = resHandler.getDebugInfo()
	Response.Write("<br/>debugInfo:" & debugInfo & "<br/>")
End If
%>