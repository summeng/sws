<!--#include file="class/alipay_notify.asp"-->
<html>
<head>
	<META http-equiv=Content-Type content="text/html; charset=gb2312">
<%
' ���ܣ�֧����ҳ����תͬ��֪ͨҳ��
' �汾��3.3
' ���ڣ�2012-07-17
' ˵����
' ���´���ֻ��Ϊ�˷����̻����Զ��ṩ���������룬�̻����Ը����Լ���վ����Ҫ�����ռ����ĵ���д,����һ��Ҫʹ�øô��롣
' �ô������ѧϰ���о�֧�����ӿ�ʹ�ã�ֻ���ṩһ���ο���
	
' //////////////ҳ�湦��˵��//////////////
' ��ҳ����ڱ������Բ���
' �ɷ���HTML������ҳ��Ĵ��롢�̻�ҵ���߼��������
' ��ҳ�����ʹ��ASP�������ߵ��ԣ�Ҳ����ʹ��д�ı�����LogResult���е��ԣ��ú����ѱ�Ĭ�Ϲرգ���alipay_notify.asp�еĺ���VerifyReturn
' ���ýӿڵ�ע��������û����ô����ɾ������
'////////////////////////////////////////
%>
<%
'����ó�֪ͨ��֤���
Set objNotify = New AlipayNotify
sVerifyResult = objNotify.VerifyReturn()

If sVerifyResult Then	'��֤�ɹ�
	'*********************************************************************
	'������������̻���ҵ���߼��������
	
	'�������������ҵ���߼�����д�������´�������ο�������
    '��ȡ֧������֪ͨ���ز������ɲο������ĵ���ҳ����תͬ��֪ͨ�����б�

	'�̻�������
	out_trade_no = Request.QueryString("out_trade_no")

	'֧�������׺�
	trade_no = Request.QueryString("trade_no")

	'����״̬
	trade_status = Request.QueryString("trade_status")

	
	If Request.QueryString("trade_status") = "WAIT_SELLER_SEND_GOODS" Then
	'�ж��Ƿ����̻���վ���Ѿ����������֪ͨ���صĴ���
		'���û������������ôִ���̻���ҵ�����
		'���������������ô��ִ���̻���ҵ�����
	ElseIf request.QueryString("trade_status") = "TRADE_FINISHED" Then
		'�жϸñʶ����Ƿ����̻���վ���Ѿ���������
			'���û�������������ݶ����ţ�out_trade_no�����̻���վ�Ķ���ϵͳ�в鵽�ñʶ�������ϸ����ִ���̻���ҵ�����
			'���������������ִ���̻���ҵ�����
	Else
		Response.Write "trade_status="&Request.QueryString("trade_status")
	End If

	Response.Write "<center>��֤�ɹ������������ɡ�</center><br>"

	'�������������ҵ���߼�����д�������ϴ�������ο�������
'��д���ݿ�֧��״̬=========================================
if request.QueryString("trade_status") = "WAIT_SELLER_SEND_GOODS" then
order_no= request.QueryString("out_trade_no")	'��ȡ�����ţ�֧����ϵͳ�������������alipayto.asp���������̼���վ�����ģ�
total_fee= request.QueryString("price")		'��ȡ�ܽ��

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
	'*********************************************************************
else '��֤ʧ��
    '��Ҫ���ԣ��뿴alipay_notify.aspҳ���VerifyReturn�������ȶ�sign��mysign��ֵ�Ƿ���ȣ����߼��responseTxt��û�з���true
    response.Write "<center>��֤ʧ��</center>"
end if
%>
<title>֧������׼˫�ӿ�</title>
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
    <td class="font_content" align="left"><%if request.QueryString("trade_status") = "WAIT_SELLER_SEND_GOODS" or request.QueryString("trade_status") = "TRADE_FINISHED" Then
	response.Write "֧���ɹ�"
	else
    response.Write request.QueryString("trade_status")
	End if%></td>
  </tr>
</table>
</body>
</html>
