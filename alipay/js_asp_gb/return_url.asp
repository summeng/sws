<!--#include file="alipay_config.asp"-->
<!--#include file="class/alipay_notify.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<%
	'���ܣ���������ת��ҳ�棨ҳ����תͬ��֪ͨҳ�棩
	'�汾��3.1
	'���ڣ�2010-10-29
	'˵����
	'���´���ֻ��Ϊ�˷����̻����Զ��ṩ���������룬�̻����Ը����Լ���վ����Ҫ�����ռ����ĵ���д,����һ��Ҫʹ�øô��롣
	'�ô������ѧϰ���о�֧�����ӿ�ʹ�ã�ֻ���ṩһ���ο���
	
''''''''ҳ�湦��˵��''''''''''''''''
'��ҳ����ڱ������Բ���
'��ҳ�������ҳ����תͬ��֪ͨҳ�桱������֧����������ͬ�����ã��ɵ�����֧����ɺ����ʾ��Ϣҳ���硰����ĳĳĳ���������ٽ����֧���ɹ�����
'�ɷ���HTML������ҳ��Ĵ���Ͷ���������ɺ�����ݿ���³������
'��ҳ�����ʹ��ASP�������ߵ��ԣ�Ҳ����ʹ��д�ı�����log_result���е��ԣ��ú����ѱ�Ĭ�Ϲرգ���alipay_notify.asp�еĺ���return_verify
'TRADE_FINISHED(��ʾ�����Ѿ��ɹ�������Ϊ��ͨ��ʱ���ʵĽ���״̬�ɹ���ʶ);
'TRADE_SUCCESS(��ʾ�����Ѿ��ɹ�������Ϊ�߼���ʱ���ʵĽ���״̬�ɹ���ʶ);
''''''''''''''''''''''''''''''''''''
%>
<%Dim verify_result,subject,order_no,total_fee
'����ó�֪ͨ��֤���
verify_result = return_verify()

if verify_result then	'��֤�ɹ�
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'������������̻���ҵ���߼��������

'��д���ݿ�֧��״̬=========================================
if request.QueryString("trade_status") = "TRADE_FINISHED" or request.QueryString("trade_status") = "TRADE_SUCCESS" then
order_no= request.QueryString("out_trade_no")	'��ȡ�����ţ�֧����ϵͳ�������������alipayto.asp���������̼���վ�����ģ�
total_fee= request.QueryString("total_fee")		'��ȡ�ܽ��

Dim username,aliname
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_order] where ddhm='"&order_no&"' and cash="&total_fee&"" 
rs.open sql,conn,1,3
username=rs("username")
aliname=rs("aliname")
rs("zfqr")=1'ͬʱ�����û���������״̬
rs("zffs")=1'ͬʱ�����û��������ʽ
rs("admincl")=1'ͬʱ�����û��������״̬
rs("ddshsj")=now()
rs("Payment")="֧����"
rs.update
Call CloseRs(rs)

set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_Financial] "
rs.open sql,conn,1,3
rs.addnew
rs("username")=username
rs("aliname")=aliname
rs("ywje")=total_fee
rs("ywlx")="����"
rs("czbz")=""&order_no&" �����������"
rs("czz")="����֧��"
rs("admincl")=1
rs("addTime")=now()
rs.update
Call CloseRs(rs)

conn.execute "UPDATE DNJY_user SET Amount=Amount+"&total_fee&" WHERE username='"&username&"'" 'ͬʱ���û������ʻ����
Call CloseDB(conn)
End if
'��д���ݿ�֧��״̬end===================================


	
	'�������������ҵ���߼�����д�������´�������ο�������
    '��ȡ֧������֪ͨ���ز������ɲο������ĵ���ҳ����תͬ��֪ͨ�����б�
    order_no		= request.QueryString("out_trade_no")	'��ȡ������
    total_fee		= request.QueryString("total_fee")		'��ȡ�ܽ��
	
	if request.QueryString("trade_status") = "TRADE_FINISHED" or request.QueryString("trade_status") = "TRADE_SUCCESS" then
		'�жϸñʶ����Ƿ����̻���վ���Ѿ����������ɲο������ɽ̡̳��С�3.4�������ݴ�����
			'���û�������������ݶ����ţ�out_trade_no�����̻���վ�Ķ���ϵͳ�в鵽�ñʶ�������ϸ����ִ���̻���ҵ�����
			'���������������ִ���̻���ҵ�����
	else
		response.Write "trade_status="&request.QueryString("trade_status")
	end if
	'�������������ҵ���߼�����д�������ϴ�������ο�������
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
else '��֤ʧ��
    '��Ҫ���ԣ��뿴alipay_notify.aspҳ���return_verify�������ȶ�sign��mysign��ֵ�Ƿ���ȣ����߼��responseTxt��û�з���true
    response.Write "fail"
end if
%>
<title>֧������ʱ֧��</title>
<style type="text/css">
            .font_content{
                font-family:"����";
                font-size:14px;
                color:#FF6600;
            }
            .font_title{
                font-family:"����";
                font-size:16px;
                color:#FF0000;
                font-weight:bold;
            }
            table{
                border: 1px solid #CCCCCC;
            }
        </style>
</head>
<body>
<table align="center" width="700" cellpadding="5" cellspacing="0">
  <tr>
    <td align="center" class="font_title" colspan="2">֪ͨ����</td>
  </tr>
  <tr>
    <td class="font_content" align="right">֧�������׺ţ�</td>
    <td class="font_content" align="left"><%=request.QueryString("trade_no")%></td>
  </tr>
  <tr>
    <td class="font_content" align="right">�����ţ�</td>
    <td class="font_content" align="left"><%=request.QueryString("out_trade_no")%></td>
  </tr>
  <tr>
    <td class="font_content" align="right">�����ܽ�</td>
    <td class="font_content" align="left"><%=request.QueryString("total_fee")%></td>
  </tr>
  <tr>
    <td class="font_content" align="right">��Ʒ���⣺</td>
    <td class="font_content" align="left"><%=request.QueryString("subject")%></td>
  </tr>
  <tr>
    <td class="font_content" align="right">��Ʒ������</td>
    <td class="font_content" align="left"><%=request.QueryString("body")%></td>
  </tr>
  <tr>
    <td class="font_content" align="right">����˺ţ�</td>
    <td class="font_content" align="left"><%=request.QueryString("buyer_email")%></td>
  </tr>
  <tr>
    <td class="font_content" align="right">����״̬��</td>
    <td class="font_content" align="left"><%if request.QueryString("trade_status") = "TRADE_FINISHED" or request.QueryString("trade_status") = "TRADE_SUCCESS" Then
	response.Write "֧���ɹ�"
	else
    response.Write request.QueryString("trade_status")
	End if%></td>
  </tr>
</table>
</body>
</html>
