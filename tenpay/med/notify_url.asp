<!--#include file="./tenpay_config.asp"-->
<!--#include file="./classes/PayResponseHandler.asp"-->
<!--#include file="./util/tenpay_util.asp"-->
<%Dim xmlDoc,obj,node
'---------------------------------------------------------
'�Ƹ�ͨ˫�ӿڴ���ص�ʾ�����̻����մ�ʾ�����п�������
'---------------------------------------------------------

log_result("�����̨֪ͨҳ�棡��������")

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
	'����notify_id��֤�������
	Dim reqHandler1
	Set reqHandler1 = new PayRequestHandler
	'�̻����յ���̨֪ͨ�����֪ͨID��Ƹ�ͨ������֤ȷ�ϣ����ú�̨ϵͳ���ý���ģʽ
	reqHandler1.setGateUrl("https://gw.tenpay.com/gateway/simpleverifynotifyid.xml")
	'��ʼ��
	reqHandler1.init()
	'������Կ
	reqHandler1.setKey(key)
	'�����̻���
	reqHandler1.setParameter "partner", partner
	'�̻�������		
	reqHandler1.setParameter "notify_id", notify_id
	'����������XMLHTTP�������
	Dim httpClient
	set httpClient = CreateObject("Msxml2.ServerXMLHTTP.3.0")
	'����֤Url����
	httpClient.Open "GET",reqHandler1.getRequestURL(),False
	'�ύ����
	httpClient.Send
	'�½�������XMLDOM�ĵ���������
	Set xmlDoc = server.CreateObject("Microsoft.XMLDOM")
	'�������󷵻ص�XML�ĵ�
	xmlDoc.load(httpClient.responseXML)
	'��ȡ�ĵ���Ԫ��
	Set obj =  xmlDoc.selectSingleNode("root")
	'����root�������ӽڵ㣬��ȡ���صļ�ֵ��
	For Each node in obj.childnodes
		reqHandler1.setParameter node.nodename, node.text
	Next
	'�Է���ֵ�����ж�
	If reqHandler1.getParameter("retcode") = "0"  Then
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
		'�ж�ǩ�������
		If "1" = trade_mode Then '��ʱ����
			If "0" = trade_state Then
			'----------------------
			'��ʱ���ʴ���ҵ��ʼ
			'-----------------------
			'�������ݿ��߼�
			'ע�⽻�׵���Ҫ�ظ�����
			'ע���жϷ��ؽ��
			'-----------------------
			'��ʱ���ʴ���ҵ�����
			'-----------------------
			'���Ƹ�ͨϵͳ���ͳɹ���Ϣ�����Ƹ�ͨϵͳ�յ��˽�����ڽ��к���֪ͨ
			log_result("��ʱ���˺�̨֪ͨ�ɹ�")
			Response.Write("success")
		    '����ɹ�
			Else  
				log_result("��ʱ���˺�̨֪ͨʧ��")
				Response.Write("��ʱ����֧��ʧ��")
			End If
		Else
			If trade_mode ="2" Then   '�н鵣��
				'-----------------------------
				'�н鵣������ҵ��ʼ
				'------------------------------
						
				'�������ݿ��߼�
				'ע�⽻�׵���Ҫ�ظ�����
				'ע���жϷ��ؽ��
'��д���ݿ�֧��״̬=============================================================
total_fee=total_fee/100
Dim username,aliname
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_order] where ddhm='"&out_trade_no&"'" '����
rs.open sql,conn,1,3
username=rs("username")
aliname=rs("aliname")
total_fee=rs("cash")
If rs("zfqr")=1 Then
'���Ƹ�ͨϵͳ���ͳɹ���Ϣ�����Ƹ�ͨϵͳ�յ��˽�����ڽ��к���֪ͨ
log_result("��ʱ���˺�̨֪ͨ�ɹ�")
Response.Write("success")
'����ɹ�
Call CloseRs(rs)
Response.end
End if
rs("zfqr")=1'ͬʱ�����û���������״̬
rs("zffs")=1'ͬʱ�����û��������ʽ
rs("admincl")=0'ͬʱ�����û��������״̬
rs("ddshsj")=now()
rs("Payment")="�Ƹ�ͨ����"
rs.update
Call CloseRs(rs)

set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_Financial] "'��Ա����
rs.open sql,conn,1,3
rs.addnew
rs("username")=username
rs("aliname")=aliname
rs("ywje")=total_fee
rs("ywlx")="����"
rs("czbz")=""&out_trade_no&" �����������"
rs("czz")="����֧��"
rs("admincl")=0
rs("addTime")=now()
rs.update
Call CloseRs(rs)
conn.execute "UPDATE DNJY_user SET Amount=Amount+"&total_fee&" WHERE username='"&username&"'" 'ͬʱ���û������ʻ����
Call CloseDB(conn)
'��д���ݿ�֧��״̬end===================================
				
				Select Case trade_state
					Case "0":	'����ɹ�
					
					Case "1":	'���״���

					Case "2":	'�ջ��ַ��д���

					Case "4":	'���ҷ����ɹ�

					Case "5":	'����ջ�ȷ�ϣ����׳ɹ�

					Case "6":	'���׹رգ�δ��ɳ�ʱ�ر�

					Case "7":	'�޸Ľ��׼۸�ɹ�

					Case "8":	'��ҷ����˿�

					Case "9":	'�˿�ɹ�

					Case "10":	'�˿�ر�

					Case else:	'error
						'nothing to do
				End Select
				'------------------------------
				'�н鵣������ҵ�����
				'------------------------------
				log_result("�н鵣����̨֪ͨ�ɹ���trade_state=" & trade_state)	
				'���Ƹ�ͨϵͳ���ͳɹ���Ϣ���Ƹ�ͨϵͳ�յ��˽�����ٽ��к���֪ͨ
				Response.Write("success")
			Else
				Response.Write("��̨����ͨ��ʧ��trade_mode����")
				'�п�����Ϊ����ԭ�������Ѿ�������δ�յ�Ӧ��
			End if
		End If
	Else
		Response.Write("Notifyid��֤ʧ��")
	End If
Else'ǩ��ʧ��
	Response.Write("ǩ����֤ʧ��")

	Dim debugInfo
	debugInfo = resHandler.getDebugInfo()
	Response.Write("<br/>debugInfo:" & debugInfo & "<br/>")

End If

%>